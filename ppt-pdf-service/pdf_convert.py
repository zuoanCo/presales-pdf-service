"""pptx -> PDF 转换器，支持两种后端，按环境自动选择：

1. powerpoint —— Windows + 安装了 Microsoft Office，COM 自动化导出。
   保真度 100%（就是 PowerPoint 自己的渲染），仅可用于有 Office 的机器。
2. libreoffice —— 跨平台（Linux/Docker/Windows/macOS），调用 soffice 无头模式。
   保真度接近 100%，前提是服务器上安装了模板用到的字体；
   复杂动画/特殊效果可能有个别差异。

选择顺序：环境变量 PDF_CONVERTER = powerpoint | libreoffice | auto（默认 auto，
先探测 PowerPoint，再用 LibreOffice）。

在独立子进程中运行（服务内通过 subprocess 调用），原因：
1. COM 要求线程初始化且 PowerPoint 不是线程安全的；
2. 子进程崩溃不会拖垮 Web 服务。

用法：python pdf_convert.py <input.pptx> <output.pdf>
"""
import os
import shutil
import subprocess
import sys


def convert_powerpoint(input_path: str, output_path: str) -> None:
    import time

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        # Office 启动慢或正忙时会返回 RPC_E_SERVERCALL_RETRYLATER (-2147417846)，
        # 稍等重试即可恢复
        last_err = None
        for _ in range(20):
            try:
                app = win32com.client.DispatchEx("PowerPoint.Application")
                break
            except pythoncom.com_error as e:
                last_err = e
                time.sleep(3)
        if app is None:
            raise RuntimeError(f"无法启动 PowerPoint（重试 20 次均失败）: {last_err}")
        # Open(FileName, ReadOnly, Untitled, WithWindow)
        # WithWindow=False 让转换在后台进行，不弹出窗口
        pres = app.Presentations.Open(input_path, True, False, False)
        # 32 = ppSaveAsPDF
        pres.SaveAs(output_path, 32)
    finally:
        if pres is not None:
            pres.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()


def convert_libreoffice(input_path: str, output_path: str) -> None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        for candidate in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            "/usr/bin/soffice",
            "/opt/libreoffice/program/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ):
            if os.path.exists(candidate):
                soffice = candidate
                break
    if soffice is None:
        raise RuntimeError("未找到 LibreOffice（soffice），请先安装")

    out_dir = os.path.dirname(output_path)
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, input_path],
        capture_output=True,
        text=True,
        timeout=300,
    )
    produced = os.path.join(out_dir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf")
    if result.returncode != 0 or not os.path.exists(produced):
        raise RuntimeError(f"LibreOffice 转换失败: {result.stderr or result.stdout}")
    if os.path.abspath(produced) != os.path.abspath(output_path):
        os.replace(produced, output_path)


def detect_backend() -> str:
    forced = os.environ.get("PDF_CONVERTER", "auto").lower()
    if forced in ("powerpoint", "libreoffice"):
        return forced
    if sys.platform == "win32":
        try:
            import win32com.client  # noqa: F401

            return "powerpoint"
        except ImportError:
            pass
    return "libreoffice"


def convert(input_path: str, output_path: str) -> str:
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    backend = detect_backend()
    if backend == "powerpoint":
        convert_powerpoint(input_path, output_path)
    else:
        convert_libreoffice(input_path, output_path)
    return backend


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: python pdf_convert.py <input.pptx> <output.pdf>", file=sys.stderr)
        sys.exit(2)
    used = convert(sys.argv[1], sys.argv[2])
    print(f"OK ({used})")
