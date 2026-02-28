# -*- coding: utf-8 -*-
"""
文档脱敏工具 v4.1
依赖：pip install python-docx python-pptx openpyxl tkinterdnd2
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import uuid, json, os, re, threading

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_OK = True
except ImportError:
    DND_OK = False

try:
    from docx import Document as DocxDoc
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

try:
    from pptx import Presentation
    PPTX_OK = True
except ImportError:
    PPTX_OK = False

try:
    import openpyxl
    XLSX_OK = True
except ImportError:
    XLSX_OK = False

MAPPING_SUFFIX = "_mask_mapping.json"
SUPPORTED_EXT  = {".docx", ".pptx", ".xlsx", ".txt"}

AI_PROMPT = (
    "【重要指令】本文档中形如 [MASK_XXXXXXXX] 的标记是敏感信息占位符，"
    "由脱敏系统自动生成。处理本文档时必须遵守：\n"
    "1. 严禁修改、翻译、解释或删除任何 [MASK_XXXXXXXX] 标记；\n"
    "2. 将每个标记视为不可分割的词语，原样保留在输出中；\n"
    "3. 仅对标记以外的正常文字执行你的任务。\n"
    "违反上述规则将导致解密失败。"
)

# ── 配色（Apple 风格）──
BG      = "#F2F2F7"
CARD    = "#FFFFFF"
BORDER  = "#D1D1D6"
T1      = "#1C1C1E"
T2      = "#6C6C70"
ORANGE  = "#FF9500"
BLUE    = "#007AFF"
GREEN   = "#34C759"
PURPLE  = "#5856D6"
RED     = "#FF3B30"
FONT    = "Microsoft YaHei"


# ==============================================================================
# 核心引擎
# ==============================================================================

class Scrubber:

    def __init__(self):
        self.mapping  = {}   # word -> mask
        self.reverse  = {}   # mask -> word
        self.json_paths = [] # 已保存的 JSON 路径列表

    def build(self, words):
        for w in words:
            w = w.strip()
            if w and w not in self.mapping:
                m = "[MASK_" + uuid.uuid4().hex[:8].upper() + "]"
                self.mapping[w] = m
                self.reverse[m] = w

    def _enc(self, text):
        """
        加密时跳过已有掩码段，
        且使用单次正则替换，避免二次污染掩码内容。
        """

        # 匹配已有 MASK
        mask_pat = re.compile(r'\[MASK_[0-9A-F]{8}\]')

        # 敏感词为空直接返回
        if not self.mapping:
            return text

        # 构造敏感词正则（按长度降序避免短词抢长词）
        words = sorted(self.mapping.keys(), key=len, reverse=True)
        word_pat = re.compile("|".join(re.escape(w) for w in words))

        parts = mask_pat.split(text)
        masks = mask_pat.findall(text)

        safe_parts = []

        for part in parts:
            # 在纯文本片段中做一次性替换
            def repl(match):
                return self.mapping.get(match.group(0), match.group(0))

            new_part = word_pat.sub(repl, part)
            safe_parts.append(new_part)

        # 拼接回去
        out = safe_parts[0]
        for mk, p in zip(masks, safe_parts[1:]):
            out += mk + p

        return out

    def _dec(self, text):
        pat = re.compile(r'\[MASK_[0-9A-F]{8}\]')
        return pat.sub(lambda m: self.reverse.get(m.group(0), m.group(0)), text)

    def save(self, out_path):
        d    = os.path.dirname(out_path) or "."
        base = os.path.splitext(os.path.basename(out_path))[0]
        mp   = os.path.join(d, base + MAPPING_SUFFIX)
        with open(mp, "w", encoding="utf-8") as f:
            json.dump({"mapping": self.mapping, "reverse": self.reverse},
                      f, ensure_ascii=False, indent=2)
        self.json_paths.append(mp)
        return mp

    def load(self, path):
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        self.mapping.update(d.get("mapping", {}))
        self.reverse.update(d.get("reverse", d.get("reverse_mapping", {})))

    def auto_load(self, doc_path):
        d    = os.path.dirname(doc_path) or "."
        base = os.path.splitext(os.path.basename(doc_path))[0]
        mp   = os.path.join(d, base + MAPPING_SUFFIX)
        if os.path.exists(mp):
            self.load(mp)
            return True, mp
        return False, mp

    # ── 文档处理 ──

    def _handle_docx(self, src, dst, fn):
        if not DOCX_OK:
            raise ImportError("请安装 python-docx")
        doc = DocxDoc(src)
        def fix(p):
            full = p.text
            new  = fn(full)
            if full == new: return
            for r in p.runs: r.text = fn(r.text)
            if p.text != new:
                for i, r in enumerate(p.runs):
                    r.text = new if i == 0 else ""
        for p in doc.paragraphs: fix(p)
        for t in doc.tables:
            for row in t.rows:
                for c in row.cells:
                    for p in c.paragraphs: fix(p)
        doc.save(dst)

    def _handle_pptx(self, src, dst, fn):
        if not PPTX_OK:
            raise ImportError("请安装 python-pptx")
        prs = Presentation(src)
        def fix_tf(tf):
            for p in tf.paragraphs:
                full = "".join(r.text for r in p.runs)
                new  = fn(full)
                if full == new: continue
                for r in p.runs: r.text = fn(r.text)
                if "".join(r.text for r in p.runs) != new:
                    for i, r in enumerate(p.runs):
                        r.text = new if i == 0 else ""
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame: fix_tf(shape.text_frame)
                if shape.has_table:
                    for row in shape.table.rows:
                        for c in row.cells: fix_tf(c.text_frame)
        prs.save(dst)

    def _handle_xlsx(self, src, dst, fn):
        if not XLSX_OK:
            raise ImportError("请安装 openpyxl")
        wb = openpyxl.load_workbook(src)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str):
                        c.value = fn(c.value)
        wb.save(dst)

    def _handle_txt(self, src, dst, fn):
        content = ""
        for enc in ("utf-8", "utf-8-sig", "gbk"):
            try:
                with open(src, "r", encoding=enc) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        with open(dst, "w", encoding="utf-8") as f:
            f.write(fn(content))

    def process(self, src, words, mode):
        if not os.path.exists(src):
            raise FileNotFoundError("文件不存在：" + src)
        ext  = os.path.splitext(src)[1].lower()
        d    = os.path.dirname(src) or "."
        base = os.path.splitext(os.path.basename(src))[0]
        dst  = os.path.join(d, base + ("_加密" if mode == "enc" else "_解密") + ext)

        fn = self._enc if mode == "enc" else self._dec

        H = {".docx": self._handle_docx, ".pptx": self._handle_pptx,
             ".xlsx": self._handle_xlsx,  ".txt":  self._handle_txt}
        if ext in H:
            H[ext](src, dst, fn)
        elif ext in {".doc", ".ppt", ".xls"}:
            raise ValueError(ext + " 为旧版格式，请另存为新格式后处理")
        else:
            raise ValueError("不支持的格式：" + ext)

        mp = self.save(dst) if mode == "enc" else None
        return dst, mp

    def batch(self, files, words, mode, cb=None):
        if mode == "enc":
            self.build(words)
        results = []
        for i, fp in enumerate(files):
            fname = os.path.basename(fp)
            try:
                out, mp = self.process(fp, words, mode)
                results.append((fp, out, None, mp))
                if cb: cb(i+1, len(files), fname, "完成")
            except Exception as e:
                results.append((fp, None, str(e), None))
                if cb: cb(i+1, len(files), fname, "失败")
        return results


# ==============================================================================
# 文件列表组件（纯 pack 布局）
# ==============================================================================

class FileList(tk.Frame):
    """
    文件拖放列表，严格使用 pack 布局，避免与外部 grid 冲突
    """

    def __init__(self, master, accent, mode, scrubber,
                 on_json=None, **kw):
        super().__init__(master, bg=CARD, **kw)
        self.accent   = accent
        self.mode     = mode        # "enc" | "dec"
        self.scrubber = scrubber
        self.on_json  = on_json
        self.files    = []
        self._build()
        if DND_OK:
            self._bind_dnd()

    def _build(self):
        # 拖放提示
        hint_bg = self._tint(self.accent, 0.92)
        self.hint = tk.Label(
            self, bg=hint_bg, fg=self.accent, cursor="hand2",
            font=(FONT, 9),
            text=("将文件拖入此处，或点击下方按钮选择"
                  if self.mode == "enc"
                  else "将文件或 JSON 映射拖入此处")
        )
        self.hint.pack(fill="x", ipady=10)
        self.hint.bind("<Button-1>", lambda e: self._pick())

        # 列表区（用 Frame 包裹 Listbox + Scrollbar，全用 pack）
        list_frame = tk.Frame(self, bg=CARD)
        list_frame.pack(fill="both", expand=True, pady=4)

        sb = tk.Scrollbar(list_frame)
        sb.pack(side="right", fill="y")

        self.lb = tk.Listbox(
            list_frame, font=(FONT, 9),
            yscrollcommand=sb.set,
            selectmode=tk.EXTENDED,
            relief="flat", bd=0,
            bg="#F8F8FA", fg=T1,
            selectbackground=self.accent,
            selectforeground="white",
            activestyle="none"
        )
        self.lb.pack(side="left", fill="both", expand=True, padx=(4, 0))
        sb.config(command=self.lb.yview)

        self.lb.bind("<Delete>",    self._del)
        self.lb.bind("<BackSpace>", self._del)
        self.lb.bind("<Button-3>",  self._ctx)

        # 按钮行（全用 pack）
        br = tk.Frame(self, bg=CARD)
        br.pack(fill="x", pady=(0, 4))

        self._mkbtn(br, "添加文件",   self.accent,  self._pick).pack(side="left", padx=(4, 3))
        self._mkbtn(br, "添加文件夹", self.accent,  self._pick_dir).pack(side="left", padx=(0, 3))
        self._mkbtn(br, "清空",       "#AEAEB2",    self._clear).pack(side="left")

        if self.mode == "dec":
            self._mkbtn(br, "载入JSON映射", PURPLE, self._pick_json).pack(side="right", padx=4)

        self.cnt = tk.Label(br, text="0 个", font=(FONT, 8),
                            bg=CARD, fg=T2)
        self.cnt.pack(side="right", padx=6)

    def _mkbtn(self, p, txt, col, cmd):
        return tk.Button(
            p, text=txt, font=(FONT, 8),
            bg=col, fg="white",
            activebackground=self._tint(col, 0.8),
            activeforeground="white",
            relief="flat", bd=0,
            padx=8, pady=4,
            cursor="hand2",
            command=cmd
        )

    @staticmethod
    def _tint(hex_c, factor):
        h = hex_c.lstrip("#")
        r = int(int(h[0:2], 16) * factor + 255 * (1-factor))
        g = int(int(h[2:4], 16) * factor + 255 * (1-factor))
        b = int(int(h[4:6], 16) * factor + 255 * (1-factor))
        return "#{:02X}{:02X}{:02X}".format(
            min(r, 255), min(g, 255), min(b, 255))

    # ── DND ──
    def _bind_dnd(self):
        for w in (self, self.hint, self.lb):
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", self._drop)

    def _drop(self, ev):
        raw   = ev.data
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        flat  = [p[0] or p[1] for p in paths]
        for p in flat:
            p = p.strip().strip('"')
            if not p:
                continue
            if p.lower().endswith(".json") and self.mode == "dec":
                self._load_json(p)
            elif os.path.isdir(p):
                self._scan_dir(p)
            else:
                self._add(p)
        self._refresh()

    # ── 文件操作 ──
    def _pick(self):
        ps = filedialog.askopenfilenames(
            title="选择文件",
            filetypes=[("支持的文件", "*.docx *.pptx *.xlsx *.txt"),
                       ("所有文件", "*.*")]
        )
        for p in ps: self._add(p)
        self._refresh()

    def _pick_dir(self):
        d = filedialog.askdirectory(title="选择文件夹")
        if d:
            self._scan_dir(d)
            self._refresh()

    def _pick_json(self):
        p = filedialog.askopenfilename(
            title="选择 JSON 映射文件",
            filetypes=[("JSON", "*.json"), ("所有文件", "*.*")]
        )
        if p: self._load_json(p)

    def _load_json(self, path):
        try:
            self.scrubber.load(path)
            n = len(self.scrubber.mapping)
            if self.on_json:
                self.on_json(n, path)
        except Exception as e:
            messagebox.showerror("载入失败", str(e))

    def _scan_dir(self, d):
        for fn in os.listdir(d):
            if os.path.splitext(fn)[1].lower() in SUPPORTED_EXT:
                self._add(os.path.join(d, fn))

    def _add(self, path):
        if not path: return
        if os.path.splitext(path)[1].lower() not in SUPPORTED_EXT: return
        if path in self.files: return
        self.files.append(path)
        self.lb.insert(tk.END, "  " + os.path.basename(path))

    def _del(self, ev=None):
        for i in list(self.lb.curselection())[::-1]:
            self.lb.delete(i)
            self.files.pop(i)
        self._refresh()

    def _ctx(self, ev):
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="删除选中", command=self._del)
        m.add_command(label="清空列表", command=self._clear)
        m.tk_popup(ev.x_root, ev.y_root)

    def _clear(self):
        self.files.clear()
        self.lb.delete(0, tk.END)
        self._refresh()

    def _refresh(self):
        self.cnt.config(text=str(len(self.files)) + " 个")


# ==============================================================================
# 主界面（纯 pack 布局）
# ==============================================================================

class App:

    def __init__(self, root):
        self.root     = root
        self.root.title("文档脱敏")
        self.root.configure(bg=BG)
        self.root.minsize(600, 580)
        self.scrubber = Scrubber()
        self._build()
        self._check_deps()

    # ── 依赖检查 ──
    def _check_deps(self):
        miss = []
        if not DOCX_OK: miss.append("python-docx")
        if not PPTX_OK: miss.append("python-pptx")
        if not XLSX_OK: miss.append("openpyxl")
        if not DND_OK:  miss.append("tkinterdnd2（拖拽）")
        if miss:
            messagebox.showwarning(
                "缺少依赖",
                "以下库未安装，对应功能不可用：\n\n" +
                "\n".join("  · " + m for m in miss) +
                "\n\n安装：\npip install python-docx python-pptx openpyxl tkinterdnd2"
            )

    # ── 构建界面（全部 pack）──
    def _build(self):
        wrap = tk.Frame(self.root, bg=BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=20)

        # ── 标题行 ──
        row0 = tk.Frame(wrap, bg=BG)
        row0.pack(fill="x", pady=(0, 14))

        tk.Label(row0, text="文档脱敏", font=(FONT, 18, "bold"),
                 bg=BG, fg=T1).pack(side="left")
        tk.Label(row0, text="本地处理  ·  数据不上传",
                 font=(FONT, 9), bg=BG, fg=T2).pack(side="left", padx=12, pady=4)

        self.badge = tk.Label(row0, text="映射：空",
                              font=(FONT, 8), bg="#E5E5EA", fg=T2,
                              padx=8, pady=2)
        self.badge.pack(side="right")

        # ── 敏感词卡片 ──
        wc = tk.Frame(wrap, bg=CARD, highlightbackground=BORDER,
                      highlightthickness=1)
        wc.pack(fill="x", pady=(0, 12))

        wi = tk.Frame(wc, bg=CARD)
        wi.pack(fill="x", padx=14, pady=10)

        tk.Label(wi, text="敏感词", font=(FONT, 10, "bold"),
                 bg=CARD, fg=T1).pack(anchor="w")
        tk.Label(wi, text="多个词请用逗号分隔（支持中文逗号）",
                 font=(FONT, 8), bg=CARD, fg=T2).pack(anchor="w", pady=(1, 5))

        self.word_box = tk.Text(
            wi, height=2, font=(FONT, 10),
            relief="flat", bd=0,
            bg="#F2F2F7", fg=T1,
            insertbackground=ORANGE,
            wrap=tk.WORD,
            highlightbackground=BORDER,
            highlightthickness=1
        )
        self.word_box.pack(fill="x", ipady=5)
        self._ph = "例：中国，1，2，3，0，政府"
        self.word_box.insert("1.0", self._ph)
        self.word_box.config(fg=T2)
        self.word_box.bind("<FocusIn>",  self._wi)
        self.word_box.bind("<FocusOut>", self._wo)

        # ── 双列文件区（用两个并排 Frame，全用 pack）──
        cols = tk.Frame(wrap, bg=BG)
        cols.pack(fill="both", expand=True, pady=(0, 12))

        # 左列：加密区
        left = tk.Frame(cols, bg=BG)
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self._section_label(left, "加密区", ORANGE)
        enc_card = tk.Frame(left, bg=CARD, highlightbackground=BORDER,
                            highlightthickness=1)
        enc_card.pack(fill="both", expand=True)

        self.enc = FileList(enc_card, ORANGE, "enc", self.scrubber)
        self.enc.pack(fill="both", expand=True, padx=8, pady=8)

        # 右列：解密区
        right = tk.Frame(cols, bg=BG)
        right.pack(side="left", fill="both", expand=True, padx=(6, 0))

        self._section_label(right, "解密区", BLUE)
        dec_card = tk.Frame(right, bg=CARD, highlightbackground=BORDER,
                            highlightthickness=1)
        dec_card.pack(fill="both", expand=True)

        self.dec = FileList(dec_card, BLUE, "dec", self.scrubber,
                            on_json=self._json_loaded)
        self.dec.pack(fill="both", expand=True, padx=8, pady=8)

        # ── 选项行 ──
        opt_row = tk.Frame(wrap, bg=BG)
        opt_row.pack(fill="x", pady=(0, 10))

        self.del_json_var = tk.BooleanVar(value=False)
        cb = tk.Checkbutton(
            opt_row,
            text="解密成功后自动删除映射文件（JSON）",
            variable=self.del_json_var,
            font=(FONT, 9), bg=BG, fg=T2,
            activebackground=BG,
            selectcolor=CARD,
            cursor="hand2"
        )
        cb.pack(side="left")

        # ── 操作按钮行 ──
        btn_row = tk.Frame(wrap, bg=BG)
        btn_row.pack(fill="x", pady=(0, 10))

        for txt, col, cmd, side in [
            ("开始加密",   ORANGE,  lambda: self._run("enc"), "left"),
            ("开始解密",   BLUE,    lambda: self._run("dec"), "left"),
            ("清空映射",   "#AEAEB2", self._clear_map,        "left"),
            ("复制AI提示词", PURPLE, self._copy_prompt,       "left"),
        ]:
            tk.Button(
                btn_row, text=txt, font=(FONT, 9, "bold"),
                bg=col, fg="white",
                activebackground=FileList._tint(col, 0.8),
                activeforeground="white",
                relief="flat", bd=0,
                padx=14, pady=7,
                cursor="hand2",
                command=cmd
            ).pack(side=side, padx=(0, 8))

        # ── 状态栏 ──
        status_bar = tk.Frame(wrap, bg=CARD, highlightbackground=BORDER,
                              highlightthickness=1)
        status_bar.pack(fill="x")

        si = tk.Frame(status_bar, bg=CARD)
        si.pack(fill="x", padx=12, pady=7)

        self.pv = tk.DoubleVar()
        style = ttk.Style()
        style.theme_use("default")
        style.configure("S.Horizontal.TProgressbar",
                        troughcolor="#E5E5EA",
                        background=ORANGE,
                        thickness=5, borderwidth=0)
        ttk.Progressbar(si, variable=self.pv, maximum=100,
                        length=180, mode="determinate",
                        style="S.Horizontal.TProgressbar").pack(side="left")

        self.status = tk.Label(si, text="就绪",
                               font=(FONT, 9), bg=CARD, fg=T2)
        self.status.pack(side="left", padx=10)

    def _section_label(self, parent, text, color):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", pady=(0, 5))
        c = tk.Canvas(f, width=8, height=8, bg=BG, highlightthickness=0)
        c.pack(side="left", pady=1)
        c.create_oval(1, 1, 7, 7, fill=color, outline="")
        tk.Label(f, text=text, font=(FONT, 10, "bold"),
                 bg=BG, fg=T1).pack(side="left", padx=5)

    # ── 输入框占位符 ──
    def _wi(self, e):
        if self.word_box.get("1.0", tk.END).strip() == self._ph:
            self.word_box.delete("1.0", tk.END)
            self.word_box.config(fg=T1)

    def _wo(self, e):
        if not self.word_box.get("1.0", tk.END).strip():
            self.word_box.insert("1.0", self._ph)
            self.word_box.config(fg=T2)

    # ── JSON 载入回调 ──
    def _json_loaded(self, n, path):
        self._upd_badge()
        self._set_status("映射已载入，共 " + str(n) + " 条", BLUE)
        messagebox.showinfo("载入成功",
                            "映射文件已加载！共 " + str(n) + " 条记录。\n"
                            "现在可将文件拖入解密区执行解密。")

    # ── 获取敏感词 ──
    def _words(self):
        raw = self.word_box.get("1.0", tk.END).strip()
        if raw == self._ph or not raw:
            return []
        return [w.strip() for w in re.split(r"[,，\n]", raw) if w.strip()]

    # ── 执行 ──
    def _run(self, mode):
        zone  = self.enc if mode == "enc" else self.dec
        files = zone.files

        if not files:
            messagebox.showwarning(
                "提示",
                "请先在" + ("加密区" if mode == "enc" else "解密区") + "添加文件。"
            )
            return

        if mode == "enc":
            words = self._words()
            if not words:
                messagebox.showwarning("提示", "请填写敏感词。")
                return
        else:
            words = []
            if not self.scrubber.reverse and files:
                ok, mp = self.scrubber.auto_load(files[0])
                if ok:
                    self._upd_badge()
                    self._set_status("自动载入映射：" + os.path.basename(mp), BLUE)
                else:
                    ans = messagebox.askokcancel(
                        "未找到映射文件",
                        "未在文件目录找到映射文件。\n\n"
                        "点击「确定」手动选择 JSON 映射文件。"
                    )
                    if ans:
                        jp = filedialog.askopenfilename(
                            title="选择映射文件",
                            filetypes=[("JSON", "*.json"), ("所有文件", "*.*")]
                        )
                        if jp:
                            try:
                                self.scrubber.load(jp)
                                self._upd_badge()
                            except Exception as ex:
                                messagebox.showerror("错误", str(ex))
                                return
                        else:
                            return
                    else:
                        return

        self.pv.set(0)
        self._set_status("处理中...", T2)

        def worker():
            def cb(cur, tot, name, st):
                self.pv.set(cur / tot * 100)
                self._set_status(
                    "(" + str(cur) + "/" + str(tot) + ") " + name + " " + st, T2)

            results = self.scrubber.batch(files, words, mode, cb)
            self.root.after(0, lambda: self._done(results, mode))

        threading.Thread(target=worker, daemon=True).start()

    def _done(self, results, mode):
        ok   = [r for r in results if r[2] is None]
        fail = [r for r in results if r[2] is not None]
        color  = ORANGE if mode == "enc" else BLUE
        action = "加密" if mode == "enc" else "解密"

        self.pv.set(100)
        self._set_status(action + " 完成：" + str(len(ok)) + " 成功，" +
                         str(len(fail)) + " 失败", color)
        self._upd_badge()

        # 解密成功后，按需删除 JSON 映射文件
        if mode == "dec" and ok and self.del_json_var.get():
            deleted = []
            for path in list(self.scrubber.json_paths):
                try:
                    if os.path.exists(path):
                        os.remove(path)
                        deleted.append(os.path.basename(path))
                except Exception:
                    pass
            if deleted:
                self._set_status("映射文件已删除：" + ", ".join(deleted), RED)

        lines = []
        if ok:
            lines.append("成功 " + str(len(ok)) + " 个：")
            for inp, out, _, mp in ok:
                lines.append("  " + os.path.basename(inp) +
                              "  →  " + os.path.basename(out))
        if fail:
            lines.append("\n失败 " + str(len(fail)) + " 个：")
            for inp, _, err, _ in fail:
                lines.append("  " + os.path.basename(inp) + "\n    " + str(err))

        if mode == "enc" and ok:
            lines.append(
                "\n建议：发给 AI 前，点击「复制AI提示词」，\n"
                "将提示词粘贴在对话最前面，AI 会保留所有占位符不变。\n"
                "本次映射已保存在内存，收到回传文件后直接拖入解密区即可。"
            )

        title = action + (" 全部成功" if not fail else " 完成（含失败项）")
        if fail:
            messagebox.showwarning(title, "\n".join(lines))
        else:
            messagebox.showinfo(title, "\n".join(lines))

    # ── 复制AI提示词 ──
    def _copy_prompt(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(AI_PROMPT)
        self.root.update()
        messagebox.showinfo(
            "已复制到剪贴板",
            "使用方法：\n\n"
            "① 在 AI 对话框最前面粘贴此提示词\n"
            "② 再粘贴 / 上传脱敏后的文件内容\n"
            "③ AI 将原样保留所有 [MASK_XXXXXXXX] 标记\n\n"
            "收到 AI 回传内容后，保存为文件，\n"
            "拖入解密区一键还原。"
        )

    # ── 清空映射 ──
    def _clear_map(self):
        self.scrubber = Scrubber()
        # 同步给两个列表区
        self.enc.scrubber = self.scrubber
        self.dec.scrubber = self.scrubber
        self._upd_badge()
        self._set_status("映射已清空", T2)
        messagebox.showinfo("已清空", "映射关系已清空，可开始新任务。")

    # ── 辅助 ──
    def _set_status(self, text, color=None):
        self.status.config(text=text, fg=color or T2)

    def _upd_badge(self):
        n = len(self.scrubber.mapping)
        if n:
            self.badge.config(text="映射：" + str(n) + " 条",
                              bg="#D1F0DA", fg=GREEN)
        else:
            self.badge.config(text="映射：空", bg="#E5E5EA", fg=T2)


# ==============================================================================
# 入口
# ==============================================================================

def main():
    root = TkinterDnD.Tk() if DND_OK else tk.Tk()
    App(root)
    footer_label = tk.Label(
    root,
    text="与 AI 对话前，请先复制提示词，并粘贴到AI对话框。",
    fg="red",
    font=("微软雅黑", 9)
    )
    footer_label.pack(side="bottom", pady=5)
    root.mainloop()


if __name__ == "__main__":
    main()