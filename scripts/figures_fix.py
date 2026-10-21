#!/usr/bin/env python3
"""Дополнительный поиск страниц-рисунков в уже распознанных (OCR) текстах.

OCR искажает подписи («Puc.3», «Рис.Т.», «Рис.2д.»), поэтому первичный поиск мог их пропустить.
Скрипт ищет подписи устойчивым шаблоном по разметке <!-- стр. N --> в text/*.md,
рендерит недостающие страницы из PDF в figures/<id>/pNNN.png и дополняет index.json.
Метки в текстах расставляет build_registers.py при следующей сборке.
"""
import csv
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(os.environ.get("OUT_DIR", ".")).resolve()
sys.path.insert(0, str(Path(__file__).parent))
from build_registers import slug  # noqa: E402
from ocr_shard import save_png  # noqa: E402

# Р/P, и/u/н, с/c — кириллица и латинские двойники; номер с типичными ошибками OCR (Т, I, l, |, З=3, б=6)
CAP = re.compile(
    r"(?:(?<![А-Яа-яA-Za-z])[РP][иuн][сc](?:унок)?|Черт(?:еж)?|Схема)\s*[.,]?\s*"
    r"([\dТTIl|ЗбО]{1,3}[а-дa-d]?)\s*[.,:]\s*([^\n]{0,110})")


def pages_with_captions(text: str):
    res = {}
    parts = re.split(r"<!-- стр\. (\d+) -->", text)
    for i in range(1, len(parts), 2):
        n, body = int(parts[i]), parts[i + 1]
        caps = [" ".join(("Рис. " + m.group(1) + ". " + m.group(2)).split()) for m in CAP.finditer(body)]
        if caps:
            res[n] = "; ".join(caps)[:240]
    return res


def main():
    rows = list(csv.DictReader(open(ROOT / "index.csv", encoding="utf-8-sig"), delimiter=";"))
    added_total, docs = 0, 0
    for r in rows:
        if not r["pdf_path"] or r["htm_path"]:
            continue
        t = ROOT / "text" / (slug(r["number"], r["id"]) + ".md")
        if not t.exists() or "Источник текста: ocr" not in t.read_text(encoding="utf-8")[:1500]:
            continue
        found = pages_with_captions(t.read_text(encoding="utf-8").split("\n---\n", 1)[-1])
        fdir = ROOT / "figures" / r["id"]
        idx_file = fdir / "index.json"
        index = json.load(open(idx_file, encoding="utf-8")) if idx_file.exists() else []
        have = {x["page"] for x in index}
        new = {n: c for n, c in found.items() if n not in have}
        for x in index:  # подпись, если раньше не нашлась
            if not x.get("caption") and x["page"] in found:
                x["caption"] = found[x["page"]]
        if not new:
            if index:
                json.dump(index, open(idx_file, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            continue
        fdir.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory() as tmp:
            for n, cap in sorted(new.items()):
                subprocess.run(["pdftoppm", "-r", "250", "-gray", "-f", str(n), "-l", str(n),
                                str(ROOT / r["pdf_path"]), f"{tmp}/p"], capture_output=True, timeout=600)
                img = next(Path(tmp).glob("p*.pgm"), None)
                if not img:
                    continue
                name = f"p{n:03d}.png"
                save_png(img, fdir / name)
                img.unlink()
                index.append({"page": n, "file": name, "caption": cap})
        index.sort(key=lambda x: x["page"])
        json.dump(index, open(idx_file, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        added_total += len(new)
        docs += 1
        print(f"{r['number']}: +{len(new)} стр. ({', '.join(map(str, sorted(new)))})", flush=True)
    print(f"Добавлено страниц-рисунков: {added_total} в {docs} документах")


if __name__ == "__main__":
    main()
