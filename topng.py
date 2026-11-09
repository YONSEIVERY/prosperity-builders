"""비상용 PDF -> 검수용 장별 PNG(1920px)."""
import os, sys
import pymupdf

pdf, out = sys.argv[1], sys.argv[2]
d = os.path.join(out, 'slides')
os.makedirs(d, exist_ok=True)
doc = pymupdf.open(pdf)
for i, page in enumerate(doc):
    z = 1920 / page.rect.width
    page.get_pixmap(matrix=pymupdf.Matrix(z, z)).save(os.path.join(d, f's{i + 1:02d}.png'))
print('pages', doc.page_count)
