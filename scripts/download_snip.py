#!/usr/bin/env python3
"""
Скачивание СНиП и пособий к СНиП с meganorm.ru.

Списки (база 2 — основная, база 1 — докачка того, чего нет в базе 2):
  list2/64514-0.htm  СНиП                      (~668)
  list2/64494-0.htm  Пособие к СНиП            (~150)
  list2/64406-0.htm  Пособие в развитие СНиП   (~4)
  list1/21-0.htm     СНиП (база 1)             (~231)
  list1/41-0.htm     Пособие к СНиП (база 1)   (~130)
  list1/129-0.htm    Пособие в развитие СНиП (база 1)

- Парсит таблицу списка: номер, название, дата введения, статус, ссылки на PDF/HTM.
- Скачивает PDF и полнотекстовые .htm, а также картинки, на которые ссылаются .htm,
  сохраняя структуру папок сайта (чтобы картинки в .htm открывались).
- Дубли между базами отбрасываются (одинаковые номер + название).
- Пишет index.csv и README.md с таблицей-реестром.
- Возобновляемый: уже скачанные файлы пропускаются.
- Ограничение по объёму (MAX_TOTAL_MB): когда репозиторий заполнен, оставшиеся
  документы записываются в remaining.csv — их можно докачать в другой репозиторий (ONLY_IDS).
"""
import csv
import html
import os
import re
import subprocess
import sys
import time
import urllib.parse
from pathlib import Path

import requests

BASE = "https://meganorm.ru"
LISTS = [  # (url, вид документа)
    (f"{BASE}/list2/64514-0.htm", "snip"),
    (f"{BASE}/list2/64494-0.htm", "posobie"),
    (f"{BASE}/list2/64406-0.htm", "posobie"),
    (f"{BASE}/list1/21-0.htm", "snip"),
    (f"{BASE}/list1/41-0.htm", "posobie"),
    (f"{BASE}/list1/129-0.htm", "posobie"),
]
ROOT = Path(os.environ.get("OUT_DIR", ".")).resolve()
FILES_DIR = ROOT / "files"
MAX_TOTAL_MB = float(os.environ.get("MAX_TOTAL_MB", "4500"))   # лимит на репозиторий
MAX_FILE_MB = 95                                               # GitHub: файл < 100 МБ
DELAY = float(os.environ.get("DELAY", "1.0"))
ASSET_DELAY = float(os.environ.get("ASSET_DELAY", "0.15"))
RUN_MINUTES = float(os.environ.get("RUN_MINUTES", "0"))        # 0 = без ограничения
T0 = time.time()
COMMIT_EVERY = int(os.environ.get("COMMIT_EVERY", "0"))
ONLY_IDS = set(filter(None, os.environ.get("ONLY_IDS", "").split(",")))
FIELDS = ["id", "number", "name", "date", "status", "kind", "base", "pdf_path", "htm_path",
          "pdf_url", "htm_url", "card_url", "note"]

S = requests.Session()
S.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Referer": BASE + "/",
})


def get(url, tries=4):
    for i in range(tries):
        try:
            r = S.get(url, timeout=120)
            if r.status_code == 200:
                return r
            if r.status_code == 404:
                return None
            print(f"  HTTP {r.status_code} {url}", flush=True)
        except requests.RequestException as e:
            print(f"  ERR {e} {url}", flush=True)
        time.sleep(5 * (i + 1))
    return None


def decode(r):
    for enc in ("utf-8", "cp1251"):
        try:
            return r.content.decode(enc)
        except UnicodeDecodeError:
            pass
    return r.content.decode("cp1251", errors="replace")


def strip_tags(s):
    return " ".join(html.unescape(re.sub(r"<[^>]+>", " ", s)).split())


def parse_list(text, list_url, kind):
    docs = []
    base = re.search(r"/list(\d)/", list_url).group(1)
    for row in re.findall(r"<tr[^>]*>(.*?)</tr>", text, flags=re.S | re.I):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, flags=re.S | re.I)
        if len(cells) < 4:
            continue
        links = [urllib.parse.urljoin(list_url, l) for l in re.findall(r'href="([^"]+)"', cells[0], flags=re.I)]
        data = [l for l in links if re.search(r"/Data\d/", l)]
        index = [l for l in links if re.search(r"/Index\d/", l)]
        if not data and not index:
            continue
        m = re.search(r"/(\d+)\.htm", index[0] if index else data[0])
        anchors = re.findall(r"<a[^>]*>(.*?)</a>", cells[0], flags=re.S)
        docs.append({
            "id": (m.group(1) if m else "") if base == "2" else f"b1-{m.group(1) if m else ''}",
            "number": strip_tags(anchors[-1]) if anchors else strip_tags(cells[0]),
            "name": strip_tags(cells[1]),
            "date": strip_tags(cells[2]),
            "status": strip_tags(cells[3]),
            "kind": kind,
            "base": base,
            "pdf_url": next((u for u in data if u.lower().endswith(".pdf")), ""),
            "htm_url": next((u for u in data if u.lower().endswith(".htm")), ""),
            "card_url": index[0] if index else "",
        })
    return docs


def list_pages(first_url):
    """Первая страница списка + пагинация вида <cat>-N.htm (если есть)."""
    m = re.match(r"(.*/)(\d+)-\d+\.htm$", first_url)
    queue, seen, pages = [first_url], set(), []
    while queue:
        u = queue.pop(0)
        if u in seen:
            continue
        seen.add(u)
        r = get(u)
        time.sleep(1)
        if not r:
            continue
        t = decode(r)
        pages.append((u, t))
        if m:
            for p in set(re.findall(rf'href="([^"]*?{m.group(2)}-\d+\.htm)"', t)):
                queue.append(urllib.parse.urljoin(u, p))
    return pages


def local_path(url):
    p = urllib.parse.urlparse(url).path.lstrip("/")
    return FILES_DIR / urllib.parse.unquote(p)


def dir_size_mb(p):
    return sum(f.stat().st_size for f in p.rglob("*") if f.is_file()) / 1048576 if p.exists() else 0


def save(url, dest, delay=None):
    if dest.exists() and dest.stat().st_size > 0:
        return "exists"
    r = get(url)
    time.sleep(DELAY if delay is None else delay)
    if r is None:
        return "fail"
    if len(r.content) > MAX_FILE_MB * 1048576:
        return "too_big"
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)
    return "ok"


def htm_assets(htm_url, dest):
    """Скачать картинки/стили, на которые ссылается полнотекстовый .htm."""
    raw = dest.read_bytes()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("cp1251", errors="replace")
    refs = set(re.findall(r'(?:src|href)="([^"#?]+\.(?:gif|png|jpe?g|bmp|css|wmf|emf|svg))"',
                          text, flags=re.I))
    for ref in refs:
        if ref.startswith(("http://", "https://")) and "meganorm.ru" not in ref:
            continue
        aurl = urllib.parse.urljoin(htm_url, ref)
        save(aurl, local_path(aurl), ASSET_DELAY)


def write_reports(docs, remaining):
    for fname, rows in (("index.csv", docs), ("remaining.csv", remaining)):
        with open(ROOT / fname, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=FIELDS, delimiter=";", extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)

    def link(p, label):
        return f"[{label}]({urllib.parse.quote(p)})" if p else ""

    n_snip = sum(d["kind"] == "snip" for d in docs)
    lines = [
        "# СНиП — Строительные нормы и правила (и пособия к СНиП)",
        "",
        "Источник: meganorm.ru — " + ", ".join(f"<{u}>" for u, _ in LISTS) + ".",
        "",
        f"Документов: **{len(docs)}** (СНиП — {n_snip}, пособия — {len(docs) - n_snip}). "
        f"Не поместилось в репозиторий: **{len(remaining)}** (см. `remaining.csv`).",
        "",
        "Полный реестр с путями — `index.csv` (разделитель `;`, открывается в Excel).",
        "",
        "> Статусы — по данным meganorm (для РФ). На территории РУз действуют ШНК/КМК; СНиП применяются "
        "там, где нет национального документа или есть прямая ссылка. Замены СНиП → СП — `registers/zamena.csv`.",
        "",
        "| № | Номер | Название | Дата введения | Статус | Файлы |",
        "|---|---|---|---|---|---|",
    ]
    for n, d in enumerate(docs, 1):
        files = " ".join(filter(None, [link(d["pdf_path"], "PDF"), link(d["htm_path"], "HTM")]))
        name = d["name"].replace("|", "/")
        lines.append(f"| {n} | {d['number']} | {name} | {d['date']} | {d['status']} | {files or '—'} |")
    readme = ROOT / "README.md"
    keep = ""
    if readme.exists():  # блок регистров (build_registers.py) сохраняем
        m = re.search(r"<!-- registers:start -->.*<!-- registers:end -->", readme.read_text(encoding="utf-8"), re.S)
        keep = m.group(0) if m else ""
    if keep:
        i = lines.index("| № | Номер | Название | Дата введения | Статус | Файлы |")
        lines[i:i] = [keep, "", "## Полный список", ""]
    readme.write_text("\n".join(lines) + "\n", encoding="utf-8")


def git_checkpoint(msg):
    if not COMMIT_EVERY:
        return
    subprocess.run(["git", "add", "-A", "files", "index.csv", "README.md"], cwd=ROOT, check=True)
    if subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=ROOT).returncode == 0:
        return
    subprocess.run(["git", "commit", "-q", "-m", msg], cwd=ROOT, check=True)
    for _ in range(3):
        if subprocess.run(["git", "push", "-q"], cwd=ROOT).returncode == 0:
            print("  -> pushed:", msg, flush=True)
            return
        subprocess.run(["git", "pull", "-q", "--rebase"], cwd=ROOT)
    print("  !! push не удался", flush=True)


def dedup_key(d):
    num = re.sub(r"\s+", "", d["number"].upper().replace("Ё", "Е"))
    name = re.sub(r"[^а-яa-z0-9]+", "", d["name"].lower().replace("ё", "е"))
    return num, name[:80]


def main():
    docs, seen_files, seen_keys = [], set(), set()
    for lu, kind in LISTS:
        got = 0
        for u, t in list_pages(lu):
            for d in parse_list(t, u, kind):
                fkey = (d["id"], d["pdf_url"], d["htm_url"])
                k = dedup_key(d)
                if fkey in seen_files or k in seen_keys:
                    continue
                seen_files.add(fkey)
                seen_keys.add(k)
                docs.append(d)
                got += 1
        print(f"Список {lu}: новых документов {got}", flush=True)
    print(f"Всего документов: {len(docs)}", flush=True)
    if not docs:
        sys.exit("Список пуст — вероятно, сайт заблокировал запрос.")

    remaining, stats, stopped = [], {}, False
    for d in docs:
        d["pdf_path"] = d["htm_path"] = d["note"] = ""
        for k in ("pdf", "htm"):  # уже скачанное в прошлых прогонах
            if d[f"{k}_url"] and local_path(d[f"{k}_url"]).exists():
                d[f"{k}_path"] = local_path(d[f"{k}_url"]).relative_to(ROOT).as_posix()
    for i, d in enumerate(docs, 1):
        if RUN_MINUTES and (time.time() - T0) / 60 > RUN_MINUTES:
            stopped = True
            break
        if ONLY_IDS and d["id"] not in ONLY_IDS:
            continue
        if dir_size_mb(FILES_DIR) > MAX_TOTAL_MB:
            d["note"] = "не скачан: лимит репозитория"
            remaining.append(d)
            continue
        print(f"[{i}/{len(docs)}] {d['number']} — {d['name'][:70]}", flush=True)
        for k in ("pdf", "htm"):
            url = d[f"{k}_url"]
            if not url:
                continue
            dest = local_path(url)
            res = save(url, dest)
            stats[res] = stats.get(res, 0) + 1
            if res in ("ok", "exists"):
                d[f"{k}_path"] = dest.relative_to(ROOT).as_posix()
                if k == "htm":
                    htm_assets(url, dest)
            else:
                d["note"] += f"{k}: {res}; "
        if COMMIT_EVERY and i % COMMIT_EVERY == 0:
            write_reports(docs, remaining)
            git_checkpoint(f"СНиП: файлы до №{i} из {len(docs)}")

    status = "complete" if not stopped and not remaining else ("incomplete" if stopped else "limit")
    write_reports(docs, remaining)
    (ROOT / ".status").write_text(status + "\n")
    print("Итог:", stats, "| объём files/, МБ:", round(dir_size_mb(FILES_DIR), 1), flush=True)


if __name__ == "__main__":
    main()
