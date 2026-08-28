"""对比原始模板 slide30 的边框定义。"""
import re
import sys
import zipfile

path = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\15944\Desktop\售前模板.pptx"
slide = sys.argv[2] if len(sys.argv) > 2 else "slide30"
z = zipfile.ZipFile(path)
xml = z.read(f"ppt/slides/{slide}.xml").decode("utf-8")
for m in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
    sp = m.group(0)
    name = re.search(r'name="([^"]+)"', sp)
    ln = re.search(r"<a:ln[ >].*?</a:ln>", sp, re.S)
    if ln:
        seg = ln.group(0)
        colors = re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', seg)
        scheme = re.findall(r'schemeClr val="(\w+)"', seg)
        nofill = "<a:noFill/>" in seg
        print(f"{name.group(1):20s} ln: srgb={colors} scheme={scheme} noFill={nofill}")
