#!/usr/bin/env python3
"""
Сборка регистров СНиП по направлениям и упрощённых текстов.

Вход:  index.csv (после download_snip.py), files/, registers/classification.csv
Выход:
  text/<номер>__<id>.md          — упрощённый текст (шапка, оглавление, чистый текст)
  registers/NN_<направление>.md  — регистр направления (таблица со ссылками)
  registers/Реестр_СНиП.xlsx     — все регистры в Excel (лист на направление)
  registers/registers.json       — машиночитаемый реестр
  AI_INDEX.md                    — компактный индекс для работы Claude (одна строка на документ)
"""
import csv
import html
import json
import os
import re
import subprocess
import sys
import urllib.parse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from snip_directions import DIRECTIONS, CODES, classify  # noqa: E402

ROOT = Path(os.environ.get("OUT_DIR", ".")).resolve()
REPO_URL = "https://github.com/ojbekhabibov-svg/SNIP/blob/main/"
OCR_DIR = Path(os.environ.get("OCR_DIR", "ocr_out"))
DIR = {d[0]: d for d in DIRECTIONS}
VALID = {"действует", "введен впервые", "взамен", "частично действует",
         "не установлен срок действия", "не определен законодательством",
         "организационно-ограниченного применения", "заменен в части",
         "действующий", "применяется", "актуализирован",
         "документ не актуализирован", "привязан к неактуализированному документу"}


# ---------------------------------------------------------------- текст из HTM
def htm_to_text(raw: bytes, htm_path: Path = None) -> str:
    try:
        t = raw.decode("utf-8")
    except UnicodeDecodeError:
        t = raw.decode("cp1251", errors="replace")
    t = re.sub(r"(?is)<(script|style|head)\b.*?</\1>", "", t)
    t = re.sub(r"(?is)<p[^>]*class=\"?MsoToc[^>]*>.*?</p>", "", t)       # оглавление Word — своё соберём
    def img(m):
        src = re.search(r'src="([^"]+)"', m.group(0), re.I)
        if not src or htm_path is None:
            return " [рис.] "
        f = (htm_path.parent / urllib.parse.unquote(src.group(1))).resolve()
        name = f.name
        kind = "ф"
        try:
            from PIL import Image
            w, h = Image.open(f).size
            ext = f.suffix.lower()
            if ext in (".jpg", ".jpeg") or (ext == ".gif" and w >= 150 and h >= 90) \
                    or (w >= 350 and h >= 250):  # PNG в этих HTM — почти всегда формулы
                kind = "РИС"
        except Exception:
            kind = "РИС" if f.suffix.lower() in (".jpg", ".jpeg") else "ф"
        return f" [{kind}:{name}] "
    t = re.sub(r"(?is)<img[^>]*>", img, t)
    t = re.sub(r"(?is)<h([1-4])[^>]*>(.*?)</h\1>",
               lambda m: "\n\n" + "#" * (int(m.group(1)) + 1) + " " + " ".join(re.sub("<[^>]+>", " ", m.group(2)).split()) + "\n\n", t)
    t = re.sub(r"(?is)</t[dh]>", " | ", t)
    t = re.sub(r"(?is)</tr>", "\n", t)
    t = re.sub(r"(?is)<br\s*/?>", "\n", t)
    t = re.sub(r"(?is)</(p|div|li|table)>", "\n", t)
    t = re.sub(r"(?s)<[^>]+>", " ", t)
    t = html.unescape(t).replace("\xa0", " ")
    out = []
    for line in t.splitlines():
        line = " ".join(line.split())
        line = re.sub(r"(\s*\|\s*)+$", "", line)
        line = re.sub(r"^\|\s*", "", line)
        if re.fullmatch(r"[|\s]*", line):
            out.append("")
            continue
        out.append(line)
    return tidy(out)


# ---------------------------------------------------------------- текст из PDF
_SPACED = re.compile(r"\b(?:[А-ЯЁA-Z0-9]{1,2} ){2,}[А-ЯЁA-Z0-9]{1,3}\b")


def fix_spaced(line: str) -> str:
    """'ПРОМ Ы Ш ЛЕННО СТИ' -> 'ПРОМЫШЛЕННОСТИ' для строк заглавными (артефакты OCR)."""
    letters = [c for c in line if c.isalpha()]
    if letters and sum(c.isupper() for c in letters) / len(letters) > 0.8:
        toks = line.split(" ")
        if sum(len(x) <= 2 for x in toks if x) > len(toks) * 0.4:
            return re.sub(r"(?<=\S) (?=\S{1,3}\b)", "", line)
    return line


def pdf_pages_text(path: Path):
    """Текстовый слой PDF через PyMuPDF (лучше pdftotext для старого OCR-слоя)."""
    try:
        import pymupdf
        with pymupdf.open(path) as d:
            return [p.get_text(sort=True) for p in d]
    except Exception:
        out = subprocess.run(["pdftotext", "-enc", "UTF-8", str(path), "-"],
                             capture_output=True, timeout=600).stdout.decode("utf-8", "replace")
        return out.split("\f")


_WORD = re.compile(r"[А-Яа-яЁё]+")
VOCAB = set()


def build_vocab(texts):
    """Словарь из чистых текстов (HTM) — для склейки разорванных слов в PDF-тексте."""
    from collections import Counter
    c = Counter()
    for t in texts:
        c.update(w.lower() for w in _WORD.findall(t))
    VOCAB.update(w for w, n in c.items() if n >= 2)


def repair_words(line: str) -> str:
    toks = line.split(" ")
    changed = True
    while changed and VOCAB:
        changed, out, i = False, [], 0
        while i < len(toks):
            if i + 1 < len(toks) and _WORD.fullmatch(toks[i] or "-"):
                m = re.match(r"([а-яё]+)(.*)", toks[i + 1] or "")
                if m:
                    a, b = toks[i], m.group(1)
                    if (a + b).lower() in VOCAB and (a.lower() not in VOCAB or b not in VOCAB):
                        out.append(a + b + m.group(2))
                        i += 2
                        changed = True
                        continue
            out.append(toks[i])
            i += 1
        toks = out
    return " ".join(toks)


def pdf_to_text(path: Path, ocr_file: Path = None, fig_pages: dict = None, method_name: str = "ocr"):
    if ocr_file and ocr_file.exists():
        pages, method = ocr_file.read_text(encoding="utf-8").split("\f"), method_name
    else:
        pages, method = pdf_pages_text(path), "pdf-text"
    lines = []
    for n, p in enumerate(pages, 1):
        p = re.sub(r"\xad\s*", "", p)
        pl = [" ".join(l.split()) for l in p.splitlines()]
        if method == "pdf-text":
            pl = [repair_words(fix_spaced(l)) for l in pl]
        pl = [l for l in pl if l and not re.fullmatch(r"(?:стр\.?|с\.)?\s*-?\s*\d{1,3}\s*-?", l, re.I)]
        if fig_pages and n in fig_pages:
            pl.insert(0, f"[РИС стр.{n}: {fig_pages[n]}]")
        if pl:
            lines.append(f"<!-- стр. {n} -->")
            lines.extend(pl)
            lines.append("")
    return tidy(join_paragraphs(lines)), method, len(pages)


def apply_fig_markers(body: str, fp: dict) -> str:
    """Снять старые метки [РИС стр.N: …] и поставить заново по figures/<id>/index.json."""
    body = re.sub(r"\[РИС стр\.\d+: [^\]]+\] ?\n?", "", body)
    return re.sub(r"<!-- стр\. (\d+) -->\n",
                  lambda m: m.group(0) + (f"[РИС стр.{m.group(1)}: {fp[int(m.group(1))]}]\n" if int(m.group(1)) in fp else ""),
                  body)


def join_paragraphs(lines):
    """Склейка строк PDF в абзацы: перенос по слогам и строки без конца предложения."""
    out = []
    for l in lines:
        if out and out[-1] and l and not l.startswith(("<!--", "#", "[РИС")) and not out[-1].startswith(("<!--", "[РИС")):
            prev = out[-1]
            starts_new = re.match(r"^(\d+(\.\d+)*\.?\s|[а-я]\)\s|-\s|—\s|•|Таблица|Рис\.|Приложение|ПРИЛОЖЕНИЕ)", l)
            if prev.endswith("-") and l[:1].islower():
                out[-1] = prev[:-1] + l
                continue
            if not starts_new and not re.search(r"[.:;!?]$", prev) and (l[:1].islower() or l[:1] in "(«\""):
                out[-1] = prev + " " + l
                continue
        out.append(l)
    return out


def tidy(lines):
    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"(\[рис\.\]\s*){2,}", "[рис.] ", text)
    text = re.sub(r" +(\[(?:ф|РИС):)", r" \1", text)
    return text.strip() + "\n"


# ---------------------------------------------------------------- рисунки
CAPTION = re.compile(r"^\s*(?:Рис\.?|Рисунок|Черт\.?|Чертеж|Схема)\s*[\dIVX№]", re.I)


def figures_htm(body: str, assets_dir: str):
    """Список рисунков из HTM-текста: (файл, подпись)."""
    lines = body.splitlines()
    figs = []
    for i, l in enumerate(lines):
        for name in re.findall(r"\[РИС:([^\]]+)\]", l):
            cap = ""
            for c in [l] + lines[i + 1:i + 4]:
                m = re.search(r"((?:Рис\.?|Рисунок|Черт\.?|Схема)\s*[\dIVX№][^|]{0,140})", c)
                if m:
                    cap = " ".join(m.group(1).split())
                    break
            figs.append((f"{assets_dir}/{name}", cap))
    return figs


def figures_pdf(doc_id: str):
    """Страницы-рисунки из OCR: figures/<id>/index.json -> [(путь, подпись, стр.)]."""
    idx = ROOT / "figures" / doc_id / "index.json"
    if not idx.exists():
        return []
    data = json.load(open(idx, encoding="utf-8"))
    return [(f"figures/{doc_id}/{x['file']}", x.get("caption", ""), x["page"]) for x in data]


def page_figs(doc_id: str):
    """Только страницы-сканы (не встроенные картинки) — для меток [РИС стр.N] в начале страницы."""
    idx = ROOT / "figures" / doc_id / "index.json"
    if not idx.exists():
        return {}
    return {x["page"]: f"figures/{doc_id}/{x['file']}" for x in json.load(open(idx, encoding="utf-8")) if not x.get("inline")}


# ---------------------------------------------------------------- оглавление
_HEAD_RX = re.compile(
    r"^(?:#{2,5} .+|(?:\d{1,2}\.){1,2}\s*[А-ЯЁ][А-ЯЁA-Z ,\-–—()]{4,}$|"
    r"(?:РАЗДЕЛ|Раздел|ГЛАВА|Глава)\s+[\dIVX]+.*|ПРИЛОЖЕНИЕ\s*\d*.*|Приложение\s+\d+.*|"
    r"\d{1,2}\.\s+[А-ЯЁ][А-ЯЁ ,\-–—]{6,})$")


def toc(text: str, limit=60):
    items = []
    for line in text.splitlines():
        s = line.strip()
        if len(s) > 160 or len(s) < 5:
            continue
        if _HEAD_RX.match(s):
            s = s.lstrip("#").strip()
            if s not in items:
                items.append(s)
        if len(items) >= limit:
            break
    return items


# ---------------------------------------------------------------- утилиты
def slug(number: str, doc_id: str) -> str:
    s = re.sub(r"[^\w.\-]+", "_", number.replace("/", "-").replace("*", "x")).strip("_") or "SNIP"
    return f"{s}__{doc_id}"


def gh(path: str) -> str:
    return REPO_URL + urllib.parse.quote(path) if path else ""


def md_link(path, label):
    return f"[{label}](../{urllib.parse.quote(path)})" if path else ""


def is_valid(status: str) -> bool:
    return status.strip().lower() in VALID


# ---------------------------------------------------------------- main
def import_figures(doc_id: str):
    """Перенести страницы-рисунки из результатов OCR в figures/<id>/."""
    src = ROOT / OCR_DIR / "figs" / doc_id
    if not src.exists():
        return
    import shutil
    dst = ROOT / "figures" / doc_id
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def _znorm(s):
    s = s.upper().replace("—", "-").replace("–", "-")
    s = re.sub(r"^\s*(?:ПОСОБИЕ\s+(?:К|В\s+РАЗВИТИЕ)\s+)?(?:СНИП|СН|СП)[\s-]*", "", s)
    m = re.match(r"((?:[IVX]+-)?\d+(?:[.\-]\d+)*)", s.replace(" ", ""))
    return m.group(1) if m else ""


def apply_zamena(rows):
    """registers/zamena.csv (old;new;uz;source): СНиП → актуализированный СП (РФ) и аналог РУз (ШНК/КМК).

    new — заменяющий документ в РФ (статус meganorm часто устарел); uz — действующий в РУз аналог.
    """
    zp = ROOT / "registers" / "zamena.csv"
    if not zp.exists():
        return
    by_num = {}
    for r in rows:
        if r.get("kind") == "snip":
            by_num.setdefault(_znorm(r["number"]), []).append(r)
    n_marked = 0
    for z in csv.DictReader(open(zp, encoding="utf-8-sig"), delimiter=";"):
        key = _znorm(z["old"])
        if not key:
            continue
        for r in by_num.get(key, []):
            if z.get("new") and "заменён" not in r["status"]:
                r["status"] = f"{r['status']}; заменён {z['new']}".strip("; ")
                n_marked += 1
            if z.get("uz"):
                r["uz"] = z["uz"]
            r["zamena_src"] = z.get("source", "")
    print("отмечено по zamena.csv:", n_marked, flush=True)


def main():
    rows = [dict(r, kind=r.get("kind") or "snip") for r in csv.DictReader(open(ROOT / "index.csv", encoding="utf-8-sig"), delimiter=";")]
    extra = ROOT / "index_extra.csv"   # документы из архива пользователя
    if extra.exists():
        rows += list(csv.DictReader(open(extra, encoding="utf-8-sig"), delimiter=";"))
    apply_zamena(rows)
    manual = {}
    cpath = ROOT / "registers" / "classification.csv"
    if cpath.exists():
        for r in csv.DictReader(open(cpath, encoding="utf-8-sig"), delimiter=";"):
            manual[(r["id"], r["number"])] = [c for c in r["directions"].split(",") if c in CODES]

    (ROOT / "text").mkdir(exist_ok=True)
    (ROOT / "registers").mkdir(exist_ok=True)
    build_vocab(htm_to_text((ROOT / r["htm_path"]).read_bytes()) for r in rows
                if r.get("htm_path") and (ROOT / r["htm_path"]).exists())
    print("словарь:", len(VOCAB), flush=True)
    docs = []
    for n, r in enumerate(rows, 1):
        if r.get("directions"):
            dirs = [c for c in r["directions"].split(",") if c in CODES]
        else:
            dirs = manual.get((r["id"], r["number"])) or classify(r["number"], r["name"])
        d = dict(r, n=n, directions=dirs, text_path="", text_src="", pages="", text_kb=0, toc=[], figures=0)
        src_htm = ROOT / r["htm_path"] if r.get("htm_path") else None
        src_pdf = ROOT / r["pdf_path"] if r.get("pdf_path") else None
        tfile = ROOT / "text" / (slug(r["number"], r["id"]) + ".md")
        body, method = "", ""
        user_txt = ROOT / "user_text" / f"{r['id']}.txt"
        has_new_ocr = (ROOT / OCR_DIR / f"{r['id']}.txt").exists() and not r.get("htm_path")
        if user_txt.exists() and src_pdf and src_pdf.exists():
            body, method, pages = pdf_to_text(src_pdf, user_txt, page_figs(r["id"]), "pdf-print")
            d["pages"] = pages
        elif tfile.exists() and os.environ.get("REBUILD_TEXT") != "1" and not has_new_ocr:
            body = tfile.read_text(encoding="utf-8").split("\n---\n", 1)[-1]
            m = re.search(r"^Источник текста: (\S+)", tfile.read_text(encoding="utf-8"), re.M)
            method = m.group(1) if m else ""
            body = ""  # пересобираем шапку и рисунки из источника, если он есть
            if method == "htm" and src_htm and src_htm.exists():
                body = htm_to_text(src_htm.read_bytes(), src_htm)
            elif src_pdf and src_pdf.exists():
                fp = {pg: path for path, cap, pg in figures_pdf(r["id"])}
                cached = ROOT / OCR_DIR / f"{r['id']}.txt"
                if method == "ocr" and not cached.exists():
                    body = apply_fig_markers(tfile.read_text(encoding="utf-8").split("\n---\n", 1)[-1], fp)
                else:
                    body, method, _ = pdf_to_text(src_pdf, cached, fp)
        elif src_htm and src_htm.exists():
            body, method = htm_to_text(src_htm.read_bytes(), src_htm), "htm"
        elif src_pdf and src_pdf.exists():
            import_figures(r["id"])
            fp = {pg: path for path, cap, pg in figures_pdf(r["id"])}
            body, method, pages = pdf_to_text(src_pdf, (ROOT / OCR_DIR / f"{r['id']}.txt"), fp)
            d["pages"] = pages
        if body.strip():
            d["toc"] = toc(body)
            if method == "htm":
                figs = [(p, c, "") for p, c in figures_htm(body, Path(r["htm_path"]).with_suffix(".files").as_posix())]
            else:
                figs = figures_pdf(r["id"])
            d["figures"] = len(figs)
            head = [
                f"# {r['number']} — {r['name']}" + (" (пособие)" if r.get("kind") == "posobie" and "особие" not in r["number"] else ""),
                "",
                f"Статус: {r['status']}" + (f" | Аналог РУз: {r['uz']}" if r.get("uz") else "") + f" | Дата введения: {r['date']} | id: {r['id']}",
                f"Направления: {', '.join(DIR[c][2] for c in dirs)}",
                f"Источник текста: {method} | Оригиналы: "
                + " ".join(filter(None, [r.get("pdf_path", ""), r.get("htm_path", "")])),
                *([f"Поступление: {r['note']}"] if r.get("note", "").startswith("архив") else []),
                "Примечание: упрощённая текстовая версия; формулы, рисунки и сложные таблицы смотреть в оригинале.",
                "",
                "## Оглавление (автоматически)",
                "",
            ] + [f"- {x}" for x in d["toc"]] + [""]
            if figs:
                head += ["## Рисунки и схемы", "",
                         ("В тексте метки [РИС:файл] — рисунок, [ф:файл] — формула картинкой; файлы в папке "
                          f"`{Path(r['htm_path']).with_suffix('.files').as_posix()}/`." if method == "htm" else
                          f"В тексте метки [РИС:figures/{r['id']}/…] — рисунок на своём месте, [ф:…] — формула/мелкая картинка; "
                          "[РИС стр.N: путь] — страница-скан с рисунком." if method == "pdf-print" else
                          "Страницы-сканы с рисунками; в тексте метки [РИС стр.N: путь]."), ""]
                head += [f"- `{p}`" + (f" стр.{pg}" if pg else "") + (f" — {c}" if c else "") for p, c, pg in figs]
                head += [""]
            head += ["---", ""]
            tfile.write_text("\n".join(head) + body, encoding="utf-8")
            d["text_path"] = tfile.relative_to(ROOT).as_posix()
            d["text_src"] = method
            d["text_kb"] = round(tfile.stat().st_size / 1024)
        docs.append(d)
        if n % 50 == 0:
            print(f"текст: {n}/{len(rows)}", flush=True)

    write_md_registers(docs)
    write_xlsx(docs)
    write_ai_index(docs)
    json.dump([{k: d.get(k, "") for k in ("id", "number", "name", "date", "status", "directions", "kind",
                                   "pdf_path", "htm_path", "text_path", "text_src", "toc")} for d in docs],
              open(ROOT / "registers" / "registers.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    update_readme(docs)
    print("Готово. Документов:", len(docs), "| с текстом:", sum(1 for d in docs if d["text_path"]))


def write_md_registers(docs):
    for code, fname, title, ppr in DIRECTIONS:
        sel = [d for d in docs if code in d["directions"]]
        sel.sort(key=lambda d: (not is_valid(d["status"]), d["number"]))
        lines = [f"# {fname.split('_', 1)[0]}. {title}", "",
                 f"Документов: **{len(sel)}**, из них действующих: **{sum(is_valid(d['status']) for d in sel)}**."
                 + (" Направление важно для ППР." if ppr else ""), "",
                 "Сначала действующие, затем заменённые/недействующие. «заменён СП …» — статус для РФ; "
                 "в РУз сначала проверять ШНК/КМК, при отсутствии — применять СНиП.", "",
                 "| № | Номер | Название | Дата | Статус | Файлы | Текст |", "|---|---|---|---|---|---|---|"]
        for i, d in enumerate(sel, 1):
            files = " ".join(filter(None, [md_link(d["pdf_path"], "PDF"), md_link(d["htm_path"], "HTM")]))
            st = d["status"] if is_valid(d["status"]) else f"~~{d['status']}~~"
            lines.append(f"| {i} | {d['number']} | {d['name'].replace('|', '/')} | {d['date']} | {st} | "
                         f"{files or '—'} | {md_link(d['text_path'], 'TXT') or '—'} |")
        (ROOT / "registers" / f"{fname}.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    idx = ["# Регистры СНиП по направлениям", "", "| Код | Направление | Документов | Действующих | ППР |", "|---|---|---|---|---|"]
    for code, fname, title, ppr in DIRECTIONS:
        sel = [d for d in docs if code in d["directions"]]
        idx.append(f"| {code} | [{title}]({urllib.parse.quote(fname)}.md) | {len(sel)} | "
                   f"{sum(is_valid(d['status']) for d in sel)} | {'да' if ppr else ''} |")
    idx += ["", "Документ может входить в несколько направлений. Распределение — `classification.csv` "
            "(колонки id;number;directions, коды через запятую); правки вносятся туда, затем пересборка workflow «Регистры СНиП».",
            "", "Excel-версия всех регистров: `Реестр_СНиП.xlsx`. Замены СНиП → СП / аналоги РУз: `zamena.csv`."]
    (ROOT / "registers" / "README.md").write_text("\n".join(idx) + "\n", encoding="utf-8")


def write_xlsx(docs):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hfill = PatternFill("solid", fgColor="CCFFFF")
    ok_fill = PatternFill("solid", fgColor="E2EFDA")
    old_fill = PatternFill("solid", fgColor="F2F2F2")
    link = Font(name="Calibri", size=10, color="0563C1", underline="single")
    heads = ["№", "Номер", "Наименование", "Дата введения", "Статус", "Направления", "PDF", "HTM", "Текст"]
    widths = [5, 24, 80, 13, 20, 26, 7, 7, 7]

    def sheet(ws, sel, title):
        ws.append([title])
        ws["A1"].font = Font(name="Calibri", size=12, bold=True)
        ws.append(heads)
        for c in range(1, len(heads) + 1):
            cell = ws.cell(row=2, column=c)
            cell.font = Font(name="Calibri", size=10, bold=True)
            cell.fill = hfill
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
        for i, d in enumerate(sel, 1):
            ws.append([i, d["number"], d["name"], d["date"], d["status"], ", ".join(d["directions"]),
                       "PDF" if d["pdf_path"] else "", "HTM" if d["htm_path"] else "", "TXT" if d["text_path"] else ""])
            r = ws.max_row
            fill = ok_fill if is_valid(d["status"]) else old_fill
            for c in range(1, len(heads) + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                cell.font = Font(name="Calibri", size=10, color="000000" if is_valid(d["status"]) else "808080")
                cell.alignment = Alignment(vertical="top", wrap_text=c in (3, 6))
                if c == 5:
                    cell.fill = fill
            for c, key in ((7, "pdf_path"), (8, "htm_path"), (9, "text_path")):
                if d[key]:
                    ws.cell(row=r, column=c).hyperlink = gh(d[key])
                    ws.cell(row=r, column=c).font = link
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:{get_column_letter(len(heads))}{ws.max_row}"

    ws = wb.active
    ws.title = "Сводка"
    ws.append(["Регистры СНиП по направлениям (источник: meganorm.ru)"])
    ws["A1"].font = Font(name="Calibri", size=12, bold=True)
    ws.append(["Код", "Направление", "Лист", "Документов", "Действующих", "Важно для ППР"])
    for c in range(1, 7):
        ws.cell(row=2, column=c).font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=2, column=c).fill = hfill
        ws.cell(row=2, column=c).border = border
    for col, w in zip("ABCDEF", (8, 70, 30, 12, 13, 13)):
        ws.column_dimensions[col].width = w
    allws = wb.create_sheet("Все СНиП")
    sheet(allws, docs, f"Все СНиП и пособия — {len(docs)} документов")
    for code, fname, title, ppr in DIRECTIONS:
        sel = sorted([d for d in docs if code in d["directions"]], key=lambda d: (not is_valid(d["status"]), d["number"]))
        name = fname[:31]
        sheet(wb.create_sheet(name), sel, f"{title} — {len(sel)} док.")
        ws.append([code, title, name, len(sel), sum(is_valid(d["status"]) for d in sel), "да" if ppr else ""])
        r = ws.max_row
        ws.cell(row=r, column=3).hyperlink = f"#'{name}'!A1"
        ws.cell(row=r, column=3).font = link
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = border
            if c != 3:
                ws.cell(row=r, column=c).font = Font(name="Calibri", size=10, bold=ppr and c == 2)
    ws.append([])
    ws.append(["Зелёный статус — действующий документ; серый текст — заменён/не действует (справочно)."])
    wb.save(ROOT / "registers" / "Реестр_СНиП.xlsx")


def write_ai_index(docs):
    lines = [
        "# AI_INDEX — компактный индекс СНиП и пособий к СНиП (для Claude)",
        "",
        "Как пользоваться (экономно):",
        "1. Найти документы по направлению/номеру/словам в этом файле (одна строка = один документ).",
        "2. Читать `text/<файл>.md`: шапка + автоглавление в начале, дальше чистый текст. Сначала оглавление, "
        "потом только нужный раздел (grep по номеру пункта/ключевому слову), а не весь файл.",
        "3. Поиск по всем текстам: `grep -ril \"ключевое слово\" text/`. PDF/HTM — только для формул, рисунков и таблиц.",
        "4. Приоритет для РУз: действующий ШНК/КМК (столбец «РУз» в registers/zamena.csv) → СНиП, если "
        "национального документа нет или на СНиП есть ссылка в РД. «ЗАМЕНЁН → СП …» — статус для РФ "
        "(актуализированная редакция); в РУз сам СНиП может продолжать применяться — проверять.",
        "Связка с ВСН: технология конкретных работ часто в ВСН (репо ojbekhabibov-svg/VSN, его AI_INDEX.md), "
        "общие требования и приёмка — в СНиП.",
        "",
        "Коды направлений: " + "; ".join(f"{c}={DIR[c][2]}" for c in CODES),
        "",
        "5. Рисунки/схемы: в шапке text-файла раздел «Рисунки и схемы» (путь + подпись); в тексте метки "
        "[РИС:файл] (HTM), [РИС стр.N: figures/<id>/pNNN.png] (скан-страница), [ф:файл] — формула картинкой. "
        "Открывать картинку (Read) только когда схема реально нужна.",
        "",
        "Формат: `номер | статус | направления | название | text-файл | КБ | источник текста | рис.` (П — пособие)",
        "",
    ]
    for code, fname, title, ppr in DIRECTIONS:
        sel = sorted([d for d in docs if code == d["directions"][0]], key=lambda d: d["number"])
        if not sel:
            continue
        lines.append(f"## {code} — {title}")
        for d in sel:
            st = "ДЕЙСТВ" if is_valid(d["status"]) else ("ЗАМЕНЁН в РФ → " + d["status"].split("заменён ", 1)[1]
                                                          if "заменён " in d["status"] else d["status"])
            if d.get("uz"):
                st += f" [РУз: {d['uz']}]"
            num = d["number"] + (" (П)" if d.get("kind") == "posobie" and "особие" not in d["number"] else "")
            lines.append(f"{num} | {st} | {','.join(d['directions'])} | {d['name']} | "
                         f"{d['text_path'].removeprefix('text/') or '—'} | {d['text_kb']} | {d['text_src'] or '—'} | {d['figures']}")
        lines.append("")
    (ROOT / "AI_INDEX.md").write_text("\n".join(lines), encoding="utf-8")


def update_readme(docs):
    p = ROOT / "README.md"
    t = p.read_text(encoding="utf-8") if p.exists() else "# СНиП\n"
    block = ["<!-- registers:start -->", "## Регистры по направлениям", "",
             "Excel: [`registers/Реестр_СНиП.xlsx`](registers/Реестр_СНиП.xlsx) · сводка: [`registers/`](registers/README.md) · "
             "упрощённые тексты: `text/` · индекс для Claude: [`AI_INDEX.md`](AI_INDEX.md)"
             + " · замены СНиП → СП и аналоги РУз: [`registers/zamena.csv`](registers/zamena.csv)", "",
             "| Направление | Документов | Действующих |", "|---|---|---|"]
    for code, fname, title, ppr in DIRECTIONS:
        sel = [d for d in docs if code in d["directions"]]
        block.append(f"| [{title}](registers/{urllib.parse.quote(fname)}.md) | {len(sel)} | {sum(is_valid(d['status']) for d in sel)} |")
    block.append("<!-- registers:end -->")
    b = "\n".join(block)
    if "<!-- registers:start -->" in t:
        t = re.sub(r"<!-- registers:start -->.*<!-- registers:end -->", b, t, flags=re.S)
    else:
        t = t.replace("\n| № | Номер |", "\n" + b + "\n\n## Полный список\n\n| № | Номер |", 1) if "\n| № | Номер |" in t else t + "\n" + b + "\n"
    p.write_text(t, encoding="utf-8")


if __name__ == "__main__":
    main()
