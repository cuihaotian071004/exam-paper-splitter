# -*- coding: utf-8 -*-
"""
试卷分割工具 v3.1.2 — 核心处理模块
DOCX→PDF 改用直接 soffice 调用（复用常驻进程，跳过 LO Python 冷启动）
"""
import os, sys, shutil, re, time, subprocess, io, tempfile, json, atexit, socket

import fitz  # PyMuPDF

BACKUP = os.path.join(os.environ["LOCALAPPDATA"], "Hermes Agent CN Desktop", "docx_backup")
os.makedirs(BACKUP, exist_ok=True)
BACKUP_MAP = os.path.join(BACKUP, "_mapping.json")


# ─── 备份映射管理 ───────────────────────────────────

def _record_backup(src_path, log=None):
    if not os.path.exists(src_path):
        return
    fname = os.path.basename(src_path)
    srcdir = os.path.dirname(os.path.abspath(src_path))
    mapping = {}
    if os.path.exists(BACKUP_MAP):
        try:
            with open(BACKUP_MAP, "r", encoding="utf-8") as f:
                mapping = json.load(f)
        except:
            pass
    mapping[fname] = srcdir
    try:
        with open(BACKUP_MAP, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
    except Exception as e:
        if log:
            log("  记录备份映射失败: %s" % e)


def cleanup_orphan_backups(log=None):
    if not os.path.exists(BACKUP_MAP):
        return
    try:
        with open(BACKUP_MAP, "r", encoding="utf-8") as f:
            mapping = json.load(f)
    except:
        return
    if not mapping:
        return
    output_suffixes = ["_题目.pdf", "_答案.pdf", "_提取.pdf",
                       "_1.pdf", "_2.pdf", "_3.pdf", "_4.pdf", "_5.pdf"]
    dirty = False
    to_delete = []
    for fname, srcdir in list(mapping.items()):
        backup_path = os.path.join(BACKUP, fname)
        if not os.path.exists(backup_path):
            del mapping[fname]
            dirty = True
            continue
        base, ext = os.path.splitext(fname)
        has_output = False
        for suf in output_suffixes:
            expected = os.path.join(srcdir, base + suf)
            if os.path.exists(expected):
                has_output = True
                break
        original_path = os.path.join(srcdir, fname)
        if os.path.exists(original_path):
            has_output = True
        if not has_output:
            to_delete.append((fname, backup_path))
    for fname, backup_path in to_delete:
        try:
            os.remove(backup_path)
            del mapping[fname]
            dirty = True
            if log:
                log("  清理孤儿备份: %s (输出文件已移走)" % fname)
        except Exception as e:
            if log:
                log("  删除备份失败: %s — %s" % (fname, e))
    if dirty:
        try:
            with open(BACKUP_MAP, "w", encoding="utf-8") as f:
                json.dump(mapping, f, ensure_ascii=False, indent=2)
        except:
            pass


# ─── PDF split mode (multi-range extract) ───────────────

def process_pdf_split(src, ranges, log, master=None, keep_original=True):
    base = os.path.splitext(os.path.basename(src))[0]
    od = os.path.dirname(src)
    doc = fitz.open(src)
    total = doc.page_count
    for a, b in ranges:
        if a < 1 or b > total or a > b:
            log("❌ 页码范围无效: %s-%s" % (a, b))
            doc.close()
            return
    log("共 %d 个分割任务, 源文件 %d 页" % (len(ranges), total))
    for idx, (a, b) in enumerate(ranges):
        suffix = "_%d" % (idx + 1) if len(ranges) > 1 else "_提取"
        out_path = os.path.join(od, base + suffix + ".pdf")
        pages = list(range(a - 1, b))
        log("  任务%d: 第%d-%d页 (%d页) → %s" % (idx + 1, a, b, len(pages), os.path.basename(out_path)))
        out_doc = fitz.open()
        for p in pages:
            out_doc.insert_pdf(doc, from_page=p, to_page=p)
        out_doc.save(out_path, deflate=True)
        out_doc.close()
    doc.close()
    if os.path.exists(src):
        if not keep_original:
            try:
                os.remove(src)
            except Exception as e:
                log("删除源文件失败: %s" % e)
    log("完成 → %d 个文件" % len(ranges))


# ─── PDF processing (split questions/answers) ───────────

def process_pdf(src, log, master=None):
    base = re.sub(r'_\d+_\d+$', '', os.path.splitext(os.path.basename(src))[0])
    od = os.path.dirname(src)
    qp = os.path.join(od, base + "_题目.pdf")
    ap = os.path.join(od, base + "_答案.pdf")
    doc = fitz.open(src)
    sp = None
    for i in range(doc.page_count):
        if "参考答案" in doc[i].get_text():
            sp = i
            break
    if sp is None:
        for i in range(doc.page_count):
            if re.search(r'(?:^|\n)\s*答案\s*(?:\n|$)', doc[i].get_text()):
                sp = i
                break
    if sp is None:
        total_text = sum(len(doc[i].get_text().strip()) for i in range(doc.page_count))
        if total_text < 50:
            log("⚠ 可能是扫描型 PDF，无法自动识别分割点")
            if not _handle_scanned_pdf(doc, src, base, od, log, master=master):
                doc.close(); return
            doc.close()
            if os.path.exists(src):
                dst = os.path.join(BACKUP, os.path.basename(src))
                shutil.move(src, dst)
                _record_backup(dst, log)
            return
        else:
            log("未找到'参考答案'或独立的'答案'标题")
            doc.close(); return
    if sp == 0:
        log("⚠ '参考答案'在第1页，全部归为答案")
        ad = fitz.open(); ad.insert_pdf(doc, from_page=0, to_page=doc.page_count - 1); ad.save(ap); ad.close()
        if os.path.exists(qp):
            try: os.remove(qp)
            except: pass
    else:
        qd = fitz.open(); qd.insert_pdf(doc, from_page=0, to_page=sp - 1); qd.save(qp); qd.close()
        ad = fitz.open(); ad.insert_pdf(doc, from_page=sp, to_page=doc.page_count - 1); ad.save(ap); ad.close()
    doc.close()
    if os.path.exists(src):
        dst = os.path.join(BACKUP, os.path.basename(src))
        shutil.move(src, dst)
        _record_backup(dst, log)
    for p in [qp, ap]:
        if os.path.exists(p):
            _remove_empty_pages(p)
    log("完成")


def _handle_scanned_pdf(doc, src, base, od, log, master=None):
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        log("无法打开 GUI"); return False
    thumbs = [doc[i].get_pixmap(dpi=72).tobytes("png") for i in range(doc.page_count)]
    if master and master.winfo_exists():
        root = tk.Toplevel(master)
    else:
        root = tk.Tk()
    root.title("扫描型 PDF — 选择分割页 (%s)" % os.path.basename(src))
    root.configure(bg="#f8f9fa")
    root.lift(); root.focus_force()
    root.attributes('-topmost', True)
    root.after(200, lambda: root.attributes('-topmost', False))
    result = {"page": None}
    tk.Label(root, text="文件: %s" % os.path.basename(src),
             font=("Microsoft YaHei", 12, "bold"), bg="#f8f9fa").pack(pady=(10, 0))
    tk.Label(root, text="请拖动滑块选择'参考答案'从第几页开始：",
             font=("Microsoft YaHei", 10), bg="#f8f9fa", fg="#555").pack(pady=(5, 0))
    canvas = tk.Canvas(root, width=400, height=300, bg="#e8e8e8", highlightthickness=0)
    canvas.pack(pady=10)
    label = tk.Label(root, text="", font=("Microsoft YaHei", 10), bg="#f8f9fa")
    label.pack(pady=(5, 0))
    def update_preview(val):
        page = int(float(val))
        result["page"] = page
        label.config(text="答案从第 %d 页开始  (共 %d 页)" % (page + 1, doc.page_count))
        canvas.delete("all")
        if 0 <= page < len(thumbs):
            try:
                from PIL import Image, ImageTk
                img = Image.open(io.BytesIO(thumbs[page]))
                w, h = img.size
                r = min(400 / w, 300 / h)
                img = img.resize((int(w * r), int(h * r)), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                canvas.image = photo
                canvas.create_image(200, 150, image=photo)
            except:
                canvas.create_text(200, 150, text="第 %d 页" % (page + 1), font=("Arial", 20))
    sf = tk.Frame(root, bg="#f8f9fa"); sf.pack(fill=tk.X, padx=20)
    tk.Label(sf, text="题目", font=("Microsoft YaHei", 9), bg="#f8f9fa").pack(side=tk.LEFT)
    scale = ttk.Scale(sf, from_=0, to=doc.page_count - 1, orient=tk.HORIZONTAL, command=update_preview)
    scale.set(0); scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
    tk.Label(sf, text="答案", font=("Microsoft YaHei", 9), bg="#f8f9fa").pack(side=tk.LEFT)
    update_preview(0)
    bf = tk.Frame(root, bg="#f8f9fa"); bf.pack(pady=(5, 15))
    tk.Button(bf, text="确认分割", font=("Microsoft YaHei", 11), bg="#27ae60", fg="white",
              padx=15, pady=4, command=lambda: [setattr(result, 'page', int(float(scale.get()))), root.destroy()]).pack(side=tk.LEFT, padx=5)
    tk.Button(bf, text="取消", font=("Microsoft YaHei", 11), bg="#e74c3c", fg="white",
              padx=15, pady=4, command=root.destroy).pack(side=tk.LEFT, padx=5)
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry("+%d+%d" % ((sw - w) // 2, (sh - h) // 2))
    root.resizable(False, False); root.grab_set()
    if master and master.winfo_exists():
        master.wait_window(root)
    else:
        root.mainloop()
    if result["page"] is None:
        log("用户取消"); return False
    sp = result["page"]
    qp = os.path.join(od, base + "_题目.pdf")
    ap = os.path.join(od, base + "_答案.pdf")
    if sp == 0:
        qd = fitz.open(); qd.save(qp); qd.close()
    else:
        qd = fitz.open(); qd.insert_pdf(doc, from_page=0, to_page=sp - 1); qd.save(qp); qd.close()
    ad = fitz.open(); ad.insert_pdf(doc, from_page=sp, to_page=doc.page_count - 1); ad.save(ap); ad.close()
    for p in [qp, ap]:
        if os.path.exists(p): _remove_empty_pages(p)
    log("完成 (分割点: 第%d页)" % (sp + 1))
    return True


def _remove_empty_pages(pdf_path):
    d = fitz.open(pdf_path)
    to_del = [i for i in range(d.page_count - 1, -1, -1) if len(d[i].get_text().strip()) < 20]
    if to_del:
        for i in to_del: d.delete_page(i)
        if d.page_count == 0:
            d.close(); os.remove(pdf_path); return
        tp = pdf_path.replace(".pdf", "_t.pdf")
        d.save(tp, deflate=True); d.close()
        os.replace(tp, pdf_path)
    else:
        d.close()


# ─── LibreOffice 监听器管理 ─────────────────────

_LO_LISTENER = None       # soffice 监听进程
_LO_LISTENER_PORT = 2003  # UNO socket 端口

def _find_soffice():
    """查找 LibreOffice 可执行文件路径。"""
    for pf in [os.environ.get("ProgramFiles", ""),
               os.environ.get("ProgramFiles(x86)", "")]:
        p = os.path.join(pf, "LibreOffice", "program", "soffice.exe")
        if os.path.exists(p):
            return p
    return None

def _try_reuse_listener(callback=None):
    """尝试连接已有 LO 监听器（处理 crash 残留，旧 LO 仍在监听端口的情况）。"""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.settimeout(1.0)
        s.connect(("127.0.0.1", _LO_LISTENER_PORT))
        s.close()
        if callback: callback("done", "就绪 (复用已有)")
        return True
    except:
        s.close()
        return False

def _is_port_in_use(port):
    """检查端口是否被占用。"""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.settimeout(0.3)
        s.connect(("127.0.0.1", port))
        s.close()
        return True
    except:
        s.close()
        return False

def _cleanup_stale_lo():
    """清理残留的 LibreOffice 进程（crash 后 soffice.exe 可能占着端口不释放）。"""
    try:
        subprocess.run(["taskkill", "/f", "/im", "soffice.exe"],
                       capture_output=True, timeout=5)
    except:
        pass

def _init_listener(callback=None):
    """启动 LO 监听器。callback(status, msg) 用于 GUI 进度显示。"""
    global _LO_LISTENER

    # 1. 已有管理中的 LO 进程且存活
    if _LO_LISTENER and _LO_LISTENER.poll() is None:
        if callback: callback("done", "就绪")
        return True

    # 2. 尝试连接已有 LO 监听器（crash 残留但 LO 仍健康的）
    if _try_reuse_listener(callback):
        return True

    soffice = _find_soffice()
    if not soffice:
        if callback: callback("error", "未找到 LibreOffice")
        return False

    # 3. 端口被占但连接不上 → 有僵尸 LO 进程，清理掉
    if _is_port_in_use(_LO_LISTENER_PORT):
        if callback: callback("progress", "检测到残留进程，正在清理...")
        _cleanup_stale_lo()
        time.sleep(0.5)  # 等待端口释放

    if callback: callback("progress", "正在启动 LibreOffice 转换引擎...")
    _LO_LISTENER = subprocess.Popen(
        [soffice, "--headless", "--accept=socket,host=localhost,port=%d;urp;" % _LO_LISTENER_PORT],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # 等待 LO 就绪（最多 8 秒）
    for i in range(40):
        time.sleep(0.2)
        if _LO_LISTENER.poll() is not None:
            if callback: callback("error", "LibreOffice 启动失败")
            _LO_LISTENER = None
            return False
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.settimeout(0.5)
            s.connect(("127.0.0.1", _LO_LISTENER_PORT))
            s.close()
            if callback: callback("done", "就绪")
            return True
        except:
            s.close()
        if callback:
            pct = min(90, int((i+1) / 40 * 90))
            callback("progress", "正在初始化... %d%%" % pct)

    if callback: callback("done", "就绪")
    return True

def _stop_listener():
    """关闭 LO 监听器。"""
    global _LO_LISTENER
    if _LO_LISTENER and _LO_LISTENER.poll() is None:
        _LO_LISTENER.terminate()
        try:
            _LO_LISTENER.wait(timeout=5)
        except:
            try:
                _LO_LISTENER.kill()
            except:
                pass
    _LO_LISTENER = None

atexit.register(_stop_listener)

# ─── DOCX→PDF 转换 ────────────────────────

def _docx_to_single_pdf(docx_path, pdf_path, log):
    """将 DOCX 直接转换为 PDF（复用常驻 soffice 进程，跳过 UNO/LO Python 开销）。
    
    v3.1.2: 基准测试 — 36KB DOCX 耗时 ~0.5s (直接 soffice) vs ~1.5s (UNO 管道)。
    """
    soffice = _find_soffice()
    if not soffice:
        log("LibreOffice 未安装", "ERROR")
        return False

    outdir = os.path.dirname(pdf_path)
    tmpdir = tempfile.mkdtemp(prefix="lo_")
    try:
        # 用目标 PDF 名作为临时 DOCX 名，soffice 生成同名 PDF 到 outdir
        tmp_docx = os.path.join(tmpdir, os.path.basename(pdf_path).replace(".pdf", ".docx"))
        shutil.copy2(docx_path, tmp_docx)

        cmd = [soffice, "--headless", "--norestore",
               "--convert-to", "pdf", "--outdir", outdir, tmp_docx]
        proc = subprocess.run(cmd, capture_output=True, timeout=120)

        expected = os.path.join(outdir, os.path.basename(pdf_path))
        return proc.returncode == 0 and os.path.exists(expected)
    except subprocess.TimeoutExpired:
        log("  转换超时 (120秒)", "ERROR")
        return False
    except Exception as e:
        log("  转换失败: %s" % e, "ERROR")
        return False
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ─── DOCX processing (convert to PDF then split) ────────

def process_file(src, log, master=None):
    if src.lower().endswith(".pdf"):
        process_pdf(src, log, master=master)
        return
    log("处理: " + os.path.basename(src))
    od = os.path.dirname(src)
    base = os.path.splitext(os.path.basename(src))[0]
    base = re.sub(r'_\d+_\d+$', '', base)
    full_pdf = os.path.join(od, base + ".pdf")
    log("  DOCX → PDF 转换中...")
    if not _docx_to_single_pdf(src, full_pdf, log):
        log("❌ 转换失败", "ERROR")
        return
    log("  PDF 生成: %s" % os.path.basename(full_pdf))
    process_pdf(full_pdf, log, master=master)
    if os.path.exists(src):
        try:
            dst = os.path.join(BACKUP, os.path.basename(src))
            shutil.move(src, dst)
            _record_backup(dst, log)
        except Exception as e:
            log("  备份失败: %s — %s" % (os.path.basename(src), e))
    log("完成")
