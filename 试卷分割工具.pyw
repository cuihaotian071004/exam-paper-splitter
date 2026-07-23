# -*- coding: utf-8 -*-
import os, sys, threading, subprocess, datetime, logging, io
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except:
    HAS_DND = False
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
BACKUP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Hermes Agent CN Desktop", "docx_backup")
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")

# 全局导入 core 模块（避免重复 importlib 解析）
import importlib.util
_core_spec = importlib.util.spec_from_file_location("core_mod",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "core.py"))
_core_mod = importlib.util.module_from_spec(_core_spec)
_core_spec.loader.exec_module(_core_mod)

def _setup_file_logger():
    os.makedirs(LOG_DIR, exist_ok=True)
    log_file = os.path.join(LOG_DIR, datetime.datetime.now().strftime("splitter_%Y%m%d.log"))
    logger = logging.getLogger("Splitter")
    logger.setLevel(logging.DEBUG)
    if not logger.handlers:
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
        logger.addHandler(fh)
    return logger

_file_logger = _setup_file_logger()

MAX_PREVIEW = 200
THUMB_W, THUMB_H = 90, 120


# ═══════════════════════════════════════════════════════════
#  Thumbnail pre-loader
# ═══════════════════════════════════════════════════════════

def preload_thumbs(pdf_path, log_cb):
    import fitz
    try:
        from PIL import Image, ImageTk
    except ImportError:
        return None
    doc = fitz.open(pdf_path)
    total = doc.page_count
    count = min(total, MAX_PREVIEW)
    log_cb("  正在生成 %d/%d 页缩略图..." % (count, total))
    if tk._default_root is None:
        r = tk.Tk(); r.withdraw()
    thumbs = []
    for i in range(count):
        pix = doc[i].get_pixmap(dpi=36)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        img = img.resize((THUMB_W, THUMB_H), Image.LANCZOS)
        thumbs.append(ImageTk.PhotoImage(img))
    doc.close()
    log_cb("  缩略图就绪 (%d 页)" % len(thumbs))
    return thumbs


# ═══════════════════════════════════════════════════════════
#  SplitDialog
# ═══════════════════════════════════════════════════════════

class SplitDialog:
    def __init__(self, master, pdf_path, thumbs, total_pages):
        self.master = master
        self.pdf_path = pdf_path
        self.thumbs = thumbs or []
        self.total = total_pages
        self.tasks = []
        self.result = None
        self.thumb_rects = []
        self.selected_page = None
        self.keep_var = tk.BooleanVar(value=True)
        self._build_ui()

    def _build_ui(self):
        self.win = tk.Toplevel(self.master)
        self.win.title("PDF 分割 - %s" % os.path.basename(self.pdf_path))
        self.win.geometry("780x700")
        self.win.minsize(680, 600)
        self.win.configure(bg="#f0f2f5")
        self.win.transient(self.master)
        self.win.lift(); self.win.focus_force()

        top = tk.Frame(self.win, bg="#ffffff", height=48)
        top.pack(fill=tk.X); top.pack_propagate(False)
        tk.Label(top, text="%s" % os.path.basename(self.pdf_path),
                 font=("Microsoft YaHei", 12, "bold"), bg="#ffffff", fg="#1a1a2e").pack(side=tk.LEFT, padx=16, pady=10)
        hint = "共 %d 页" % self.total if len(self.thumbs) >= self.total else "共 %d 页 (预览前%d页)" % (self.total, len(self.thumbs))
        tk.Label(top, text=hint, font=("Microsoft YaHei", 10), bg="#ffffff", fg="#7f8c8d").pack(side=tk.RIGHT, padx=16, pady=10)

        tk.Label(self.win, text="页面预览 (点击缩略图快速定位起止页)", font=("Microsoft YaHei", 9),
                 bg="#f0f2f5", fg="#95a5a6").pack(anchor="w", padx=16, pady=(10, 2))
        thumb_outer = tk.Frame(self.win, bg="#ffffff", highlightbackground="#e0e0e0", highlightthickness=1)
        thumb_outer.pack(fill=tk.X, padx=16, pady=(0, 8))
        self.thumb_canvas = tk.Canvas(thumb_outer, bg="#ffffff", height=140, highlightthickness=0)
        self.thumb_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.thumb_canvas.bind("<Configure>", lambda e: self._arrange_thumbs())
        self.thumb_canvas.bind("<Button-1>", self._on_thumb_click)
        self.thumb_canvas.bind("<MouseWheel>", lambda e: self.thumb_canvas.yview_scroll(int(-e.delta/120), "units"))

        tasks_header = tk.Frame(self.win, bg="#f0f2f5")
        tasks_header.pack(fill=tk.X, padx=16, pady=(6, 2))
        tk.Label(tasks_header, text="分割任务", font=("Microsoft YaHei", 11, "bold"),
                 bg="#f0f2f5", fg="#1a1a2e").pack(side=tk.LEFT)
        tk.Button(tasks_header, text="+ 添加任务", font=("Microsoft YaHei", 10),
                  command=self._add_task, bg="#e8f4fd", fg="#2980b9",
                  relief=tk.FLAT, padx=12, pady=3, cursor="hand2",
                  activebackground="#d4eafc").pack(side=tk.RIGHT)
        self.tasks_container = tk.Frame(self.win, bg="#f0f2f5")
        self.tasks_container.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        bottom = tk.Frame(self.win, bg="#ffffff", height=56)
        bottom.pack(fill=tk.X, side=tk.BOTTOM); bottom.pack_propagate(False)
        bf = tk.Frame(bottom, bg="#ffffff"); bf.pack(expand=True)
        tk.Button(bf, text="确认分割", font=("Microsoft YaHei", 11, "bold"),
                  command=self._on_confirm, bg="#3498db", fg="white",
                  relief=tk.FLAT, padx=24, pady=6, cursor="hand2").pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="取消", font=("Microsoft YaHei", 11),
                  command=self._on_cancel, bg="#ecf0f1", fg="#2c3e50",
                  relief=tk.FLAT, padx=24, pady=6, cursor="hand2").pack(side=tk.LEFT, padx=6)
        tk.Checkbutton(bottom, text="保留原文件", variable=self.keep_var,
                        font=("Microsoft YaHei", 9), bg="#ffffff", fg="#555",
                        selectcolor="#ffffff", activebackground="#ffffff").pack(side=tk.RIGHT, padx=16)

        self.win.update_idletasks()
        w, h = self.win.winfo_width(), self.win.winfo_height()
        sw, sh = self.win.winfo_screenwidth(), self.win.winfo_screenheight()
        self.win.geometry("+%d+%d" % ((sw - w) // 2, (sh - h) // 2))
        self.win.resizable(True, True)
        self.win.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._add_task()
        self.win.after(50, self._arrange_thumbs)

    def _arrange_thumbs(self):
        if not self.thumbs:
            self.thumb_canvas.delete("all")
            self.thumb_canvas.create_text(300, 70, text="(无缩略图 - 请手动输入页码)",
                                          font=("Microsoft YaHei", 10), fill="#999")
            return
        self.thumb_canvas.delete("all")
        self.thumb_rects = []
        cw = self.thumb_canvas.winfo_width()
        if cw < 50: return
        pad = 6
        cols = max(1, (cw - 10) // (THUMB_W + pad))
        x_start = max(4, (cw - cols * (THUMB_W + pad) + pad) // 2)
        for i, photo in enumerate(self.thumbs):
            if photo is None: continue
            col = i % cols; row = i // cols
            x = x_start + col * (THUMB_W + pad)
            y = 6 + row * (THUMB_H + pad + 18)
            self.thumb_canvas.create_image(x + THUMB_W // 2, y + THUMB_H // 2, image=photo)
            color = "#3498db" if self.selected_page == i else "#d5d8dc"
            self.thumb_canvas.create_rectangle(x, y, x + THUMB_W, y + THUMB_H, outline=color, width=2)
            self.thumb_canvas.create_text(x + THUMB_W // 2, y + THUMB_H + 10,
                                           text=str(i + 1), font=("Consolas", 9), fill="#555")
            self.thumb_rects.append((x, y, x + THUMB_W, y + THUMB_H, i))
        shown = len([p for p in self.thumbs if p is not None])
        rows = (shown + cols - 1) // cols
        total_h = rows * (THUMB_H + pad + 18) + 6
        self.thumb_canvas.config(height=min(total_h, 400), scrollregion=(0, 0, cw, total_h))

    def _on_thumb_click(self, event):
        for x1, y1, x2, y2, pn in self.thumb_rects:
            if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                self.selected_page = pn
                self._arrange_thumbs()
                active = getattr(self, '_active_task_idx', len(self.tasks) - 1)
                if active < 0 or active >= len(self.tasks): active = len(self.tasks) - 1
                if self.tasks:
                    sv, ev, _, _ = self.tasks[active]
                    page_1b = pn + 1
                    try:
                        if int(sv.get() or 1) > page_1b:
                            sv.set(str(page_1b)); ev.set(str(page_1b))
                        else:
                            ev.set(str(page_1b))
                    except ValueError:
                        sv.set(str(page_1b)); ev.set(str(page_1b))
                    self._update_preview(active)
                break

    def _add_task(self):
        idx = len(self.tasks) + 1
        start_var = tk.StringVar(value="1")
        end_var = tk.StringVar(value=str(min(self.total, 5)))
        card = tk.Frame(self.tasks_container, bg="#ffffff",
                        highlightbackground="#e0e0e0", highlightthickness=1)
        card.pack(fill=tk.X, pady=3, padx=2)
        tk.Label(card, text="任务 %d" % idx, font=("Microsoft YaHei", 10, "bold"),
                 bg="#ffffff", fg="#2c3e50", width=8, anchor="w").pack(side=tk.LEFT, padx=(12, 4), pady=8)
        inf = tk.Frame(card, bg="#ffffff"); inf.pack(side=tk.LEFT, pady=8)
        vcmd = (self.win.register(lambda v: v == "" or (v.isdigit() and 1 <= int(v) <= self.total)), '%P')
        def make_spin(var):
            return tk.Spinbox(inf, textvariable=var, from_=1, to=self.total,
                              width=5, font=("Consolas", 11), justify="center",
                              relief=tk.SOLID, borderwidth=1, validate="key", validatecommand=vcmd,
                              command=lambda v=var, i=idx-1: self._on_spin_change(i))
        s1 = make_spin(start_var); s1.pack(side=tk.LEFT)
        tk.Label(inf, text=" - ", font=("Consolas", 10), bg="#ffffff", fg="#999").pack(side=tk.LEFT)
        s2 = make_spin(end_var); s2.pack(side=tk.LEFT)
        tk.Label(inf, text="  页", font=("Microsoft YaHei", 9), bg="#ffffff", fg="#999").pack(side=tk.LEFT)
        preview = tk.Label(card, text="", font=("Microsoft YaHei", 9), bg="#ffffff", fg="#27ae60")
        preview.pack(side=tk.LEFT, padx=(8, 0), pady=8)
        tk.Button(card, text="X", font=("Arial", 12),
                  command=lambda i=idx-1: self._remove_task(i),
                  bg="#ffffff", fg="#e74c3c", relief=tk.FLAT,
                  cursor="hand2", activebackground="#fdecea",
                  padx=6, pady=2).pack(side=tk.RIGHT, padx=(0, 8), pady=8)
        for sp in [s1, s2]:
            sp.bind("<FocusIn>", lambda e, i=idx-1: setattr(self, '_active_task_idx', i))
        self.tasks.append((start_var, end_var, preview, card))
        self._update_preview(idx - 1)

    def _on_spin_change(self, idx): self._update_preview(idx)

    def _remove_task(self, idx):
        if len(self.tasks) <= 1: return
        _, _, _, card = self.tasks[idx]
        card.destroy(); del self.tasks[idx]
        for i, (_, _, pv, c) in enumerate(self.tasks):
            for child in c.winfo_children():
                if isinstance(child, tk.Label) and child.cget("text").startswith("任务"):
                    child.config(text="任务 %d" % (i + 1)); break
            self._update_preview(i)

    def _update_preview(self, idx):
        if idx >= len(self.tasks): return
        sv, ev, preview, _ = self.tasks[idx]
        try:
            a, b = int(sv.get()), int(ev.get())
            if a > b:
                preview.config(text="起止颠倒", fg="#e74c3c")
            else:
                pages = b - a + 1
                base = os.path.splitext(os.path.basename(self.pdf_path))[0]
                suffix = "_%d" % (idx + 1) if len(self.tasks) > 1 else "_提取"
                preview.config(text="> %d页 %s%s.pdf" % (pages, base, suffix), fg="#27ae60")
        except ValueError:
            preview.config(text="", fg="#999")

    def _on_confirm(self):
        ranges = []
        for sv, ev, _, _ in self.tasks:
            try:
                a, b = int(sv.get()), int(ev.get())
                if a < 1 or b > self.total or a > b:
                    tk.messagebox.showwarning("无效范围", "页码范围无效: %d-%d" % (a, b), parent=self.win)
                    return
                ranges.append((a, b))
            except ValueError:
                return
        if not ranges: return
        self.result = (ranges, self.keep_var.get())
        self.win.destroy()

    def _on_cancel(self):
        self.result = None
        self.win.destroy()

    def run(self):
        self.win.grab_set()
        self.master.wait_window(self.win)
        return self.result


# ═══════════════════════════════════════════════════════════
#  Main App
# ═══════════════════════════════════════════════════════════

class App(TkinterDnD.Tk if HAS_DND else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("试卷分割工具 v3.1.1")
        self.geometry("620x620")
        self.minsize(550, 520)
        self.configure(bg="#f8f9fa")
        self.mode = "qa"
        self._lo_ready = False  # LO 引擎是否就绪

        self._build_ui()
        # 主窗口直接显示，LO 引擎后台静默初始化，状态栏实时反馈
        self.progress.start(10)
        self.status.config(text="LO 引擎初始化中...", fg="#e67e22")
        threading.Thread(target=self._init_lo_bg, daemon=True).start()

    def _init_lo_bg(self):
        """后台初始化 LO 监听器，通过 after 更新状态栏。"""
        def update_status(st, msg):
            """线程安全的 UI 更新。"""
            def _upd():
                if st == "done":
                    self._lo_ready = True
                    self.progress.stop()
                    self.status.config(text="就绪，拖入文件或点击选择", fg="#27ae60")
                    self.log_msg("LO 引擎就绪")
                elif st == "error":
                    self.progress.stop()
                    self.status.config(text="备用模式: 就绪，拖入文件或点击选择", fg="#e67e22")
                    self.log_msg("LO 引擎未就绪，使用备用转换方式", "WARNING")
                elif st == "progress":
                    self.status.config(text="LO 引擎初始化中... %s" % msg, fg="#e67e22")
            try:
                self.after(0, _upd)
            except tk.TclError:
                pass

        try:
            _core_mod._init_listener(update_status)
        except Exception as e:
            update_status("error", str(e)[:50])

    def _build_ui(self):
        t1 = tk.Label(self, font=("Microsoft YaHei", 16, "bold"), bg="#f8f9fa", fg="#2c3e50")
        t1["text"] = "试卷分割工具 v3.1.1"
        t1.pack(pady=(10, 0))

        mf = tk.Frame(self, bg="#f8f9fa"); mf.pack(pady=(8, 0))
        self.btn_qa = tk.Button(mf, text="分离题目答案", font=("Microsoft YaHei", 11, "bold"),
                                 command=self.set_mode_qa, bg="#3498db", fg="white",
                                 relief=tk.FLAT, padx=14, pady=4, cursor="hand2")
        self.btn_qa.pack(side=tk.LEFT, padx=3)
        self.btn_split = tk.Button(mf, text="分割PDF", font=("Microsoft YaHei", 11),
                                    command=self.set_mode_split, bg="#ecf0f1", fg="#2c3e50",
                                    relief=tk.FLAT, padx=14, pady=4, cursor="hand2")
        self.btn_split.pack(side=tk.LEFT, padx=3)
        self.mode_hint = tk.Label(self, font=("Microsoft YaHei", 9), bg="#f8f9fa", fg="#7f8c8d")
        self.mode_hint.pack(pady=(4, 0))

        self.drop_frame = tk.Frame(self, bg="#eaf2f8", highlightbackground="#3498db", highlightthickness=2)
        self.drop_frame.pack(padx=15, pady=(4, 6), fill=tk.X)
        self.drop_label = tk.Label(self.drop_frame, font=("Microsoft YaHei", 14), bg="#eaf2f8", fg="#2c3e50", height=3)
        self.drop_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        if HAS_DND:
            for w in [self.drop_frame, self.drop_label]:
                w.drop_target_register(DND_FILES); w.dnd_bind("<<Drop>>", self.on_drop)

        btnf = tk.Frame(self, bg="#f8f9fa"); btnf.pack(pady=(0, 4))
        for txt, cmd, bg_c, fg_c, fs in [
            ("选择文件", self.browse_file, "#3498db", "white", 11),
            ("打开备份", self.open_backup, "#ecf0f1", "#2c3e50", 11),
            ("管理备份", self.open_backup_manager, "#f5b7b1", "#922b21", 11),
            ("日志", self.open_logs, "#d6eaf8", "#1a5276", 9),
            ("更新日志", self.show_changelog, "#f9e79f", "#7d6608", 9),
            ("使用说明", self.show_help, "#d5f5e3", "#1e8449", 9),
        ]:
            btn = tk.Button(btnf, font=("Microsoft YaHei", fs), command=cmd,
                           bg=bg_c, fg=fg_c, relief=tk.RAISED, padx=8, pady=2, cursor="hand2")
            btn["text"] = txt; btn.pack(side=tk.LEFT, padx=3)

        self.file_label = tk.Label(self, font=("Microsoft YaHei", 9), bg="#f8f9fa", fg="#7f8c8d")
        self.file_label.pack()

        lf = tk.Frame(self, bg="#f8f9fa"); lf.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
        tk.Label(lf, font=("Microsoft YaHei", 10, "bold"), bg="#f8f9fa", fg="#2c3e50",
                 text="运行日志").pack(anchor="w")
        self.log = scrolledtext.ScrolledText(lf, height=14, font=("Consolas", 10),
                                              bg="#1e1e1e", fg="#d4d4d4", relief=tk.FLAT,
                                              borderwidth=3, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        sf = tk.Frame(self, bg="#f8f9fa"); sf.pack(fill=tk.X, padx=15, pady=(0, 8))
        self.status = tk.Label(sf, font=("Microsoft YaHei", 9), bg="#f8f9fa", fg="#95a5a6",
                                text="就绪，拖入文件或点击选择")
        self.status.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(sf, mode="indeterminate", length=150)
        self.progress.pack(side=tk.RIGHT, padx=(10, 0))
        self.running = False

        _file_logger.info("="*50)
        _file_logger.info("试卷分割工具 v3.1.1 启动")
        _file_logger.info("="*50)
        self.set_mode_qa()
        self.update_idletasks()

    def set_mode_qa(self):
        self.mode = "qa"
        self.btn_qa.config(bg="#3498db", fg="white", font=("Microsoft YaHei", 11, "bold"))
        self.btn_split.config(bg="#ecf0f1", fg="#2c3e50", font=("Microsoft YaHei", 11))
        self.mode_hint["text"] = "拖拽 .docx / .pdf  自动分离题目和答案 > 生成 PDF"
        self.drop_label["text"] = "将 .docx 或 .pdf 文件拖到这里"

    def set_mode_split(self):
        self.mode = "split"
        self.btn_qa.config(bg="#ecf0f1", fg="#2c3e50", font=("Microsoft YaHei", 11))
        self.btn_split.config(bg="#e67e22", fg="white", font=("Microsoft YaHei", 11, "bold"))
        self.mode_hint["text"] = "拖拽 PDF > 可视化选择页码范围 > 分为多个文件"
        self.drop_label["text"] = "将 PDF 文件拖到这里"

    def on_drop(self, event):
        raw = event.data
        path = raw.strip().strip("{}").strip('"')
        if self.mode == "qa":
            if path.lower().endswith((".docx", ".pdf")):
                self.file_label["text"] = os.path.basename(path)
                self.file_label["fg"] = "#2c3e50"
                self.start_processing(path)
            else:
                self.log_msg("仅支持 .docx 和 .pdf 文件")
        else:
            if path.lower().endswith(".pdf"):
                self.file_label["text"] = os.path.basename(path)
                self.file_label["fg"] = "#2c3e50"
                self.start_processing(path)
            else:
                self.log_msg("PDF分割模式仅支持 .pdf 文件")

    def log_msg(self, msg, level="INFO"):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)
        self.update_idletasks()
        _file_logger.log(getattr(logging, level, logging.INFO), msg)

    def open_logs(self):
        subprocess.run(["explorer", LOG_DIR]) if os.path.exists(LOG_DIR) else self.log_msg("日志文件夹不存在")

    def browse_file(self):
        if self.mode == "qa":
            f = filedialog.askopenfilename(title="选择文件", filetypes=[("Word/PDF","*.docx;*.pdf"),("所有文件","*.*")])
        else:
            f = filedialog.askopenfilename(title="选择 PDF 文件", filetypes=[("PDF 文件","*.pdf"),("所有文件","*.*")])
        if f:
            self.file_label["text"] = os.path.basename(f); self.file_label["fg"] = "#2c3e50"
            self.start_processing(f)

    def open_backup(self):
        subprocess.run(["explorer", BACKUP_DIR]) if os.path.exists(BACKUP_DIR) else self.log_msg("备份文件夹不存在")

    def open_backup_manager(self):
        win = tk.Toplevel(self); win.title("备份文件管理"); win.geometry("720x500")
        win.minsize(550,350); win.configure(bg="#f8f9fa"); win.transient(self)
        tk.Label(win, text="管理备份文件", font=("Microsoft YaHei",14,"bold"), bg="#f8f9fa", fg="#2c3e50").pack(pady=(10,2))
        tk.Label(win, text="同源文件自动合并显示", font=("Microsoft YaHei",9), bg="#f8f9fa", fg="#7f8c8d").pack(pady=(0,8))
        fr = tk.Frame(win, bg="#f8f9fa"); fr.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0,5))
        tree = ttk.Treeview(fr, show="tree", height=16)
        tree["columns"] = ("#1","#2","#3","#4")
        tree.column("#0", width=280); tree.heading("#0", text="文件组")
        tree.column("#1", width=60, anchor="center"); tree.heading("#1", text="勾选")
        tree.column("#2", width=70, anchor="center"); tree.heading("#2", text="大小")
        tree.column("#3", width=70, anchor="center"); tree.heading("#3", text="文件数")
        tree.column("#4", width=130); tree.heading("#4", text="修改时间")
        vsb = ttk.Scrollbar(fr, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); vsb.pack(side=tk.RIGHT, fill=tk.Y)
        def get_group(fn):
            n = fn
            for suf in ["_题目","_答案","_提取","_1","_2","_3","_4","_5",".docx"]: n = n.replace(suf,"")
            import re as re2
            return re2.sub(r'_\\d+_\\d+$','',n)
        def refresh():
            for i in tree.get_children(): tree.delete(i)
            if not os.path.exists(BACKUP_DIR): return
            fl = [f for f in os.listdir(BACKUP_DIR) if os.path.isfile(os.path.join(BACKUP_DIR,f))]
            groups = {}
            for fn in fl:
                g = get_group(fn)
                if g not in groups: groups[g] = []
                fp = os.path.join(BACKUP_DIR,fn)
                groups[g].append((fn, os.path.getsize(fp), datetime.datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d %H:%M")))
            for g in sorted(groups.keys(), reverse=True):
                items = groups[g]; total_s = sum(x[1] for x in items)
                sz = str(round(total_s/1024,1))+"KB" if total_s<1048576 else str(round(total_s/1048576,1))+"MB"
                parent = tree.insert("","end",text=" "+g,values=("",sz,str(len(items))+"个文件",""))
                tree.item(parent, open=True)
                for fn,s,mt in items:
                    sz2 = str(s)+"B" if s<1024 else (str(round(s/1024,1))+"KB" if s<1048576 else str(round(s/1048576,1))+"MB")
                    tree.insert(parent,"end",text="  "+fn,values=("",sz2,"",mt))
            if not fl: tree.insert("","end",text="无备份文件")
        refresh()
        def toggle_check():
            item = tree.focus()
            if not item or tree.parent(item)=="": return
            v = list(tree.item(item,"values")); v[0] = "checked" if v[0]=="" else ""; tree.item(item, values=v)
        tree.bind("<ButtonRelease-1>", lambda e: toggle_check() if tree.identify_region(e.x,e.y)=="cell" and tree.identify_column(e.x)=="#1" else None)
        def get_sel():
            sel = []
            for item in tree.get_children():
                for child in tree.get_children(item):
                    if tree.item(child,"values")[0]=="checked": sel.append(tree.item(child,"text").strip())
            return sel
        def del_sel():
            sel = get_sel()
            if not sel: self.quiet_msg("提示","请先勾选"); return
            if self.quiet_askyesno("确认","删除 %d 个文件？" % len(sel)):
                for fn in sel:
                    try: os.remove(os.path.join(BACKUP_DIR,fn))
                    except: pass
                refresh(); self.quiet_msg("已删除","%d 个文件" % len(sel))
        def del_grp():
            item = tree.focus()
            if not item or tree.parent(item)!="": return
            g = tree.item(item,"text").strip()
            if self.quiet_askyesno("确认","删除 %s 组？" % g):
                for child in tree.get_children(item):
                    try: os.remove(os.path.join(BACKUP_DIR,tree.item(child,"text").strip()))
                    except: pass
                refresh(); self.quiet_msg("已删除","完成")
        def del_all():
            if self.quiet_askyesno("确认","清空全部？"):
                for fn in os.listdir(BACKUP_DIR):
                    try: os.remove(os.path.join(BACKUP_DIR,fn))
                    except: pass
                refresh(); self.quiet_msg("已清空","完成")
        bf = tk.Frame(win, bg="#f8f9fa"); bf.pack(fill=tk.X, padx=15, pady=(5,12))
        for t,c,bg in [("删除选中",del_sel,"#e74c3c"),("删除分组",del_grp,"#e67e22"),("清空全部",del_all,"#c0392b"),
                        ("刷新",refresh,"#3498db"),("打开文件夹",lambda: subprocess.run(["explorer",BACKUP_DIR]),"#2ecc71")]:
            tk.Button(bf,text=t,font=("Microsoft YaHei",10),command=c,bg=bg,fg="white",relief=tk.RAISED,padx=12,pady=3,cursor="hand2").pack(side=tk.LEFT,padx=4)
        tk.Button(bf,text="关闭",font=("Microsoft YaHei",10),command=win.destroy,bg="#ecf0f1",fg="#2c3e50",relief=tk.RAISED,padx=12,pady=3,cursor="hand2").pack(side=tk.RIGHT,padx=4)

    def quiet_askyesno(self, title, msg):
        self._ask_result = False
        w = tk.Toplevel(self); w.title(title); w.geometry("380x160"); w.minsize(300,130)
        w.configure(bg="#f8f9fa"); w.transient(self); w.grab_set()
        tk.Label(w,text=msg,font=("Microsoft YaHei",10),bg="#f8f9fa",fg="#2c3e50",wraplength=340,justify="left").pack(padx=20,pady=(20,15),fill=tk.BOTH,expand=True)
        bf2 = tk.Frame(w,bg="#f8f9fa"); bf2.pack(pady=(0,12))
        tk.Button(bf2,text="是",font=("Microsoft YaHei",10),command=lambda:[setattr(self,'_ask_result',True),w.destroy()],bg="#e74c3c",fg="white",relief=tk.RAISED,padx=15,pady=3,cursor="hand2").pack(side=tk.LEFT,padx=5)
        tk.Button(bf2,text="否",font=("Microsoft YaHei",10),command=w.destroy,bg="#ecf0f1",fg="#2c3e50",relief=tk.RAISED,padx=15,pady=3,cursor="hand2").pack(side=tk.LEFT,padx=5)
        self.wait_window(w)
        return self._ask_result

    def quiet_msg(self, title, msg):
        w = tk.Toplevel(self); w.title(title); w.geometry("400x220"); w.minsize(300,150)
        w.configure(bg="#f8f9fa"); w.transient(self); w.grab_set()
        tk.Label(w,text=msg,font=("Microsoft YaHei",10),bg="#f8f9fa",fg="#2c3e50",wraplength=360,justify="left").pack(padx=20,pady=(25,15),fill=tk.BOTH,expand=True)
        tk.Button(w,text="确定",font=("Microsoft YaHei",10),command=w.destroy,bg="#3498db",fg="white",relief=tk.RAISED,padx=20,pady=4,cursor="hand2").pack(pady=(0,15))

    def show_changelog(self):
        self.quiet_msg("更新日志", "试卷分割工具 版本历史\n---\nv3.1.1  启动流程简化 + 闪退修复 + LO 自动恢复\nv3.1.0  UNO API 管道转换，DOCX\u2192PDF 提速约40倍\nv3.0.0  架构重构: 统一PDF管道 + 备份管理器\nv2.1.0  扫描型PDF支持 + 多项修复\nv2.0.0  PDF分割模式\nv1.1.0  文件日志系统\nv1.0.0  初始版本")

    def show_help(self):
        self.quiet_msg("使用说明",
            "分离题目答案: 拖入.docx/.pdf \u2192 自动分割\n\n"
            "分割PDF: 拖入PDF \u2192 可视选择页码范围\n"
            "  \u00b7 点击缩略图快速定位起止页\n"
            "  \u00b7 可添加多个任务")

    def start_processing(self, file_path):
        if self.running: self.log_msg("处理中，请等待"); return
        self.running = True
        self.status["text"] = "处理中..."; self.status["fg"] = "#e67e22"
        self.progress.start(10)
        _file_logger.info("开始处理 [%s]: %s" % (self.mode, file_path))
        self.log_msg("="*40); self.log_msg(file_path)
        if self.mode == "qa":
            if file_path.lower().endswith(".pdf"):
                self._process_qa(file_path)
            else:
                threading.Thread(target=self._process_qa, args=(file_path,), daemon=True).start()
        else:
            if file_path.lower().endswith(".pdf"):
                thumbs = preload_thumbs(file_path, self.log_msg)
                import fitz
                doc = fitz.open(file_path)
                total = doc.page_count
                doc.close()
                dlg = SplitDialog(self, file_path, thumbs, total)
                result = dlg.run()
                if result is None:
                    self.log_msg("用户取消")
                else:
                    ranges, keep_original = result
                    self.log_msg("确认分割: %d 个任务" % len(ranges) + (" (保留原文件)" if keep_original else " (删除原文件)"))
                    _core_mod.process_pdf_split(file_path, ranges, self.log_msg, master=self, keep_original=keep_original)
            else:
                self.log_msg("仅支持 .pdf")
            self._done()

    def _process_qa(self, file_path):
        try:
            _core_mod.process_file(file_path, self.log_msg, master=self)
        except Exception as e:
            import traceback
            self.log_msg("错误: " + str(e), "ERROR")
            for line in traceback.format_exc().split("\n"):
                if line.strip(): self.log_msg("  " + line, "ERROR")
        finally:
            self.after(0, self._done)

    def _done(self):
        self.running = False; self.progress.stop()
        self.status["text"] = "完成"; self.status["fg"] = "#27ae60"
        try:
            _core_mod.cleanup_orphan_backups(self.log_msg)
        except Exception:
            pass


if __name__ == "__main__":
    _crash_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "crash.log")
    try:
        App().mainloop()
    except Exception:
        import traceback as _tb
        import datetime as _dt
        with open(_crash_log_path, "a", encoding="utf-8") as _f:
            _f.write("=" * 50 + "\n")
            _f.write(_dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
            _tb.print_exc(file=_f)
            _f.write("=" * 50 + "\n\n")
        raise
