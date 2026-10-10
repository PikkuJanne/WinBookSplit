"""Separate developer PDFium renderer, invoked with the bundled Python only."""
from hashlib import sha256
from importlib.metadata import version
import argparse
import json
from pathlib import Path
import platform

from PIL import Image
import pypdfium2 as pdfium


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--probe', action='store_true')
    parser.add_argument('--input', type=Path)
    parser.add_argument('--output-prefix', type=Path)
    args = parser.parse_args()
    if args.probe:
        print(json.dumps({'python': platform.python_version(), 'pypdfium2': version('pypdfium2'),
                          'pdfium': str(pdfium.PDFIUM_INFO), 'module_path': str(Path(pdfium.__file__).resolve())}, ensure_ascii=True))
        return 0
    if args.input is None or args.output_prefix is None:
        parser.error('--input and --output-prefix required unless --probe')
    document = pdfium.PdfDocument(str(args.input))
    records = []
    try:
        for index in range(len(document)):
            page = document[index]
            bitmap = page.render(scale=1.0, draw_annots=True)
            try:
                image = bitmap.to_pil().convert('RGBA')
                path = args.output_prefix.with_name(args.output_prefix.name + f'-{index + 1:03d}.png')
                with path.open('xb') as handle:
                    image.save(handle, format='PNG')
                records.append({'page': index + 1, 'filename': path.name, 'width': image.width, 'height': image.height,
                                'png_sha256': sha256(path.read_bytes()).hexdigest(), 'pixels_sha256': sha256(image.tobytes()).hexdigest()})
            finally:
                bitmap.close()
                page.close()
    finally:
        document.close()
    print(json.dumps({'renderer': 'PDFium', 'pages': records}, ensure_ascii=True))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
