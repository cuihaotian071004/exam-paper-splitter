# -*- coding: utf-8 -*-
"""
LibreOffice UNO 转换脚本 — 由 LO 自带 Python 执行。
通过 UNO socket 连接已运行的监听器，转换 DOCX 为 PDF。
用法: LO_PYTHON lo_convert.py <docx_path> <pdf_path> [port]
"""
import sys, os, traceback

def convert(docx_path, pdf_path, port=2003):
    """通过 UNO API 将 DOCX 转换为 PDF。"""
    import uno
    from com.sun.star.beans import PropertyValue

    # 连接运行中的 LibreOffice 监听器
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)

    ctx = resolver.resolve(
        "uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % port)
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # 打开 DOCX
    abs_docx = os.path.abspath(docx_path)
    doc_url = "file:///" + abs_docx.replace("\\", "/")
    doc = desktop.loadComponentFromURL(doc_url, "_blank", 0, ())

    # 导出 PDF
    abs_pdf = os.path.abspath(pdf_path)
    pdf_url = "file:///" + abs_pdf.replace("\\", "/")

    props = []
    p = PropertyValue()
    p.Name = "FilterName"
    p.Value = "writer_pdf_Export"
    props.append(p)

    doc.storeToURL(pdf_url, tuple(props))
    doc.close(True)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法: LO_PYTHON lo_convert.py <docx_path> <pdf_path> [port]", file=sys.stderr)
        sys.exit(1)

    docx_path = sys.argv[1]
    pdf_path = sys.argv[2]
    port = int(sys.argv[3]) if len(sys.argv) > 3 else 2003

    try:
        convert(docx_path, pdf_path, port)
        print("OK")
    except Exception as e:
        print("ERROR: %s" % e, file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)
