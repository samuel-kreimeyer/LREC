from pypdf import PdfWriter
from pathlib import Path

def make_test_pdf(path: Path, pages: int = 3):
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=72, height=72)
    with open(path, 'wb') as f:
        writer.write(f)

if __name__ == '__main__':
    make_test_pdf(Path('test_multi.pdf'), pages=3)
    print('Created test_multi.pdf')
