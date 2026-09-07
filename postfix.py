"""Post-process the pptxgenjs output.

1. Strip the stray mid-paragraph <a:pPr><a:buNone/></a:pPr> that pptxgenjs emits
   before the second run of a two-run bullet (invalid OOXML).
2. Inject the TEGNA wordmark as a vector custGeom shape (copied from the MW
   rebuild) — black in the footer of every content slide, white in the cover
   band of slide 1. pptxgenjs cannot write custGeom, so this is done here.

Usage: python postfix.py REPORT_MASTER_v0_2.pptx
"""
import re, zipfile, shutil, os, sys

src = sys.argv[1] if len(sys.argv) > 1 else "REPORT_MASTER_v0_2.pptx"
tmp = "unp2"
EMU = 914400
W_IN, H_IN, M_IN = 13.333, 7.5, 0.5

# TEGNA wordmark path (w=856615 h=161925 EMU, i.e. 0.937 x 0.177 in), from the MW deck.
TEGNA_PATH = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "tegna_path.xml"), encoding="utf8").read().strip()

def tegna_sp(shape_id, name, x_in, y_in, h_in, color):
    w_in = h_in * 856615 / 161925
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x_in*EMU)}" y="{int(y_in*EMU)}"/>'
        f'<a:ext cx="{int(w_in*EMU)}" cy="{int(h_in*EMU)}"/></a:xfrm>'
        f'{TEGNA_PATH}<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    )

shutil.rmtree(tmp, ignore_errors=True)
zipfile.ZipFile(src).extractall(tmp)
fixed = 0
injected = 0
for fn in sorted(os.listdir(f"{tmp}/ppt/slides")):
    if not fn.endswith(".xml"):
        continue
    p = f"{tmp}/ppt/slides/{fn}"
    x = open(p, encoding="utf8").read()

    # --- 1. stray pPr ---
    def fix_para(m):
        global fixed
        para = m.group(0)
        pprs = list(re.finditer(r'<a:pPr\b[^>]*/>|<a:pPr\b[^>]*>.*?</a:pPr>', para, flags=re.S))
        if len(pprs) <= 1:
            return para
        out = para
        for mm in reversed(pprs[1:]):
            out = out[:mm.start()] + out[mm.end():]
            fixed += 1
        return out
    x2 = re.sub(r'<a:p>.*?</a:p>', fix_para, x, flags=re.S)

    # --- 2. TEGNA wordmark ---
    n = int(re.search(r'slide(\d+)\.xml', fn).group(1))
    max_id = max(int(i) for i in re.findall(r'<p:cNvPr id="(\d+)"', x2))
    if n == 1:
        # cover band: white, to the left of the white PREMION png (which sits at x = W - M - 1.26, y 0.5, h 0.21)
        h = 0.21
        w = h * 856615 / 161925
        sp = tegna_sp(max_id + 1, "TegnaLogo", W_IN - M_IN - 1.26 - 0.15 - w, 0.5, h, "FFFFFF")
    else:
        sp = tegna_sp(max_id + 1, "TegnaLogo", M_IN, H_IN - 0.43, 0.177, "000000")
    x2 = x2.replace("</p:spTree>", sp + "</p:spTree>", 1)
    injected += 1

    if x2 != x:
        open(p, "w", encoding="utf8").write(x2)

os.remove(src)
cwd = os.getcwd()
os.chdir(tmp)
os.system(f"zip -Xr ../{src} . >/dev/null")
os.chdir(cwd)
shutil.rmtree(tmp, ignore_errors=True)
print("stray pPr removed:", fixed, "| TEGNA marks injected:", injected)
