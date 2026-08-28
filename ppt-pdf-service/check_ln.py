"""检查 slide30 XML 中所有形状的边框定义，找出黑框来源。"""
import re
import zipfile

z = zipfile.ZipFile("templates/售前模板.pptx")
xml = z.read("ppt/slides/slide30.xml").decode("utf-8")
for m in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
    sp = m.group(0)
    name = re.search(r'name="([^"]+)"', sp)
    ln = re.search(r"<a:ln[ >].*?</a:ln>", sp, re.S)
    if ln:
        colors = re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', ln.group(0))
        scheme = re.findall(r'schemeClr val="(\w+)"', ln.group(0))
        print(f"{name.group(1):20s} ln: srgb={colors} scheme={scheme}")
