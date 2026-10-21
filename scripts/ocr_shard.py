#!/usr/bin/env python3
"""OCR PDF-документов без .htm (их текстовый слой — старый OCR с разрывами слов).

Запуск: python ocr_shard.py <shard> <total> <out_dir>
Берёт документы, у которых есть PDF и нет HTM, с int(id) % total == shard;
пропускает те, у которых в text/ уже есть версия «Источник текста: ocr».
Результат: <out_dir>/<id>.txt — страницы, разделённые символом \\f.
"""
import concurrent.futures as cf
import csv
import os
import subprocess
import sys
import json
import re
import tempfile
import zlib
from pathlib import Path

ROOT = Path(os.environ.get("OUT_DIR", ".")).resolve()
DPI = os.environ.get("OCR_DPI", "250")


def ocr_page(img):
    r = subprocess.run(["tesseract", str(img), "-", "-l", "rus+eng", "--psm", "3"],
                       capture_output=True, timeout=900)
    return r.stdout.decode("utf-8", "replace")


CAPTION = re.compile(r"((?:Рис|Рисунок|Черт|Чертеж|Схема)[.,]?\s*\d+[.,]?\s*(?:[а-я]\)?\s*[—–-]|[А-ЯЁ])[^\n]{0,100})")


def figure_pages(pages, imgs):
    """Страницы с рисунками: есть подпись «Рис./Черт./Схема N» или мало текста при заметной заливке."""
    from PIL import Image
    lens = sorted(len(p.strip()) for p in pages if p.strip())
    median = lens[len(lens) // 2] if lens else 0
    res = []
    for n, (txt, img) in enumerate(zip(pages, imgs), 1):
        caps = [" ".join(c.split()) for c in CAPTION.findall(txt)]
        sparse = False
        if median and 3 < n < len(pages) and len(txt.strip()) < 0.35 * median:
            im = Image.open(img).convert("L")
            im.thumbnail((400, 400))
            dark = sum(1 for v in (im.get_flattened_data() if hasattr(im, 'get_flattened_data') else im.getdata()) if v < 128) / (im.size[0] * im.size[1])
            sparse = 0.02 < dark < 0.6
        if caps or sparse:
            res.append((n, "; ".join(caps)))
    return res


def save_png(img, dst):
    from PIL import Image
    im = Image.open(img).convert("L")
    if im.size[0] > 1400:
        im = im.resize((1400, int(im.size[1] * 1400 / im.size[0])), Image.LANCZOS)
    im.quantize(16).save(dst, optimize=True)


def main():
    shard, total, out = int(sys.argv[1]), int(sys.argv[2]), Path(sys.argv[3])
    out.mkdir(parents=True, exist_ok=True)
    sys.path.insert(0, str(Path(__file__).parent))
    from build_registers import slug

    rows = list(csv.DictReader(open(ROOT / "index.csv", encoding="utf-8-sig"), delimiter=";"))
    todo = []
    for r in rows:
        if os.environ.get("ONLY_ID") and r["id"] != os.environ["ONLY_ID"]:
            continue
        if not r["pdf_path"] or r["htm_path"] or zlib.crc32(r["id"].encode()) % total != shard:
            continue
        t = ROOT / "text" / (slug(r["number"], r["id"]) + ".md")
        if t.exists() and "Источник текста: ocr" in t.read_text(encoding="utf-8")[:1500]:
            continue
        if (ROOT / r["pdf_path"]).exists():
            todo.append(r)
    print(f"shard {shard}/{total}: документов {len(todo)}", flush=True)
    workers = os.cpu_count() or 2
    for i, r in enumerate(todo, 1):
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run(["pdftoppm", "-r", DPI, "-gray", str(ROOT / r["pdf_path"]), f"{tmp}/p"],
                           capture_output=True, timeout=3600)
            imgs = sorted(Path(tmp).glob("p*.pgm"))
            with cf.ThreadPoolExecutor(workers) as ex:
                pages = list(ex.map(ocr_page, imgs))
            figs = figure_pages(pages, imgs)
            if figs:
                fdir = out / "figs" / r["id"]
                fdir.mkdir(parents=True, exist_ok=True)
                index = []
                for n, cap in figs:
                    name = f"p{n:03d}.png"
                    save_png(imgs[n - 1], fdir / name)
                    index.append({"page": n, "file": name, "caption": cap})
                json.dump(index, open(fdir / "index.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        (out / f"{r['id']}.txt").write_text("\f".join(pages), encoding="utf-8")
        print(f"[{i}/{len(todo)}] {r['number']} — {len(imgs)} стр., рисунков: {len(figs)}", flush=True)


if __name__ == "__main__":
    main()
