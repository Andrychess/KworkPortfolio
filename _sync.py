"""
Статическая сборка портфолио для Git / GitHub Pages.

Источник: Works/<категория>/<плитка>/*
Скрипт генерирует превью PDF/PPTX и вшивает галерею в index.html.

  python _sync.py          # пересобрать сайт
  python _sync.py --open   # пересобрать и открыть index.html
"""

from __future__ import annotations

import argparse
import hashlib
import html
import re
import shutil
import webbrowser
from pathlib import Path

SITE = Path(__file__).resolve().parent
WORKS = SITE / "Works"
INDEX = SITE / "index.html"

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
PDF_EXT = {".pdf"}
PPTX_EXT = {".pptx"}
PREVIEW_DIRNAME = "previews"
PREVIEW_LIMIT = 5


def natural_key(name: str):
    return [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", name.lower())]


def web_path(path: Path) -> str:
    return path.relative_to(SITE).as_posix()


def ru_plural(n: int, one: str, few: str, many: str) -> str:
    n = abs(n) % 100
    if 11 <= n <= 19:
        return many
    n = n % 10
    if n == 1:
        return one
    if 2 <= n <= 4:
        return few
    return many


def child_dirs(folder: Path) -> list[Path]:
    return sorted(
        (p for p in folder.iterdir() if p.is_dir() and not p.name.startswith(".")),
        key=lambda p: natural_key(p.name),
    )


def esc(value: str) -> str:
    return html.escape(value, quote=True)


# ── previews ─────────────────────────────────────────────────────────────────

def _stamp(source: Path) -> str:
    st = source.stat()
    raw = f"{source.name}:{st.st_mtime_ns}:{st.st_size}"
    return hashlib.md5(raw.encode()).hexdigest()[:12]


def _cached_previews(out_dir: Path, source: Path) -> list[Path] | None:
    stamp = out_dir / ".stamp"
    if not stamp.exists() or stamp.read_text(encoding="utf-8").strip() != _stamp(source):
        return None
    slides = sorted(out_dir.glob("slide-*.png"), key=lambda p: natural_key(p.name))
    return slides or None


def _reset_dir(out_dir: Path):
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)


def make_pdf_previews(pdf: Path, out_dir: Path) -> list[Path]:
    cached = _cached_previews(out_dir, pdf)
    if cached is not None:
        return cached

    import fitz

    _reset_dir(out_dir)
    doc = fitz.open(pdf)
    result = []
    try:
        for i in range(min(PREVIEW_LIMIT, doc.page_count)):
            page = doc[i]
            zoom = 480 / page.rect.width
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            out = out_dir / f"slide-{i + 1}.png"
            pix.save(out)
            result.append(out)
    finally:
        doc.close()

    (out_dir / ".stamp").write_text(_stamp(pdf), encoding="utf-8")
    return result


def make_pptx_previews(pptx: Path, out_dir: Path) -> list[Path]:
    cached = _cached_previews(out_dir, pptx)
    if cached is not None:
        return cached

    try:
        import win32com.client
    except ImportError:
        print(f"  ! нужен pywin32 для превью PPTX: {pptx.name}")
        return []

    _reset_dir(out_dir)
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 1
    presentation = None
    result = []
    try:
        presentation = app.Presentations.Open(str(pptx.resolve()), WithWindow=False)
        for i in range(1, min(PREVIEW_LIMIT, presentation.Slides.Count) + 1):
            out = out_dir / f"slide-{i}.png"
            presentation.Slides(i).Export(str(out.resolve()), "PNG")
            result.append(out)
    except Exception as exc:
        print(f"  ! превью PPTX ({pptx.name}): {exc}")
        result = []
    finally:
        if presentation is not None:
            presentation.Close()
        app.Quit()

    if result:
        (out_dir / ".stamp").write_text(_stamp(pptx), encoding="utf-8")
    return result


# ── scan ─────────────────────────────────────────────────────────────────────

def list_media(tile_dir: Path):
    images, pdfs, pptxs = [], [], []
    for path in tile_dir.iterdir():
        if not path.is_file() or path.name.startswith("."):
            continue
        ext = path.suffix.lower()
        if ext in IMAGE_EXT:
            images.append(path)
        elif ext in PDF_EXT:
            pdfs.append(path)
        elif ext in PPTX_EXT:
            pptxs.append(path)
    key = lambda p: natural_key(p.name)
    images.sort(key=key)
    pdfs.sort(key=key)
    pptxs.sort(key=key)
    return images, pdfs, pptxs


def build_tile(category: str, tile_dir: Path) -> dict | None:
    images, pdfs, pptxs = list_media(tile_dir)
    if not images and not pdfs and not pptxs:
        return None

    title = tile_dir.name
    preview_root = tile_dir / PREVIEW_DIRNAME
    image_slides = [web_path(p) for p in images]
    files_count = len(images) + len(pdfs) + len(pptxs)

    if pdfs:
        doc = pdfs[0]
        slides = image_slides[:]
        try:
            slides = [web_path(p) for p in make_pdf_previews(doc, preview_root / "pdf")] + slides
        except Exception as exc:
            print(f"  ! PDF превью ({doc.name}): {exc}")
        bits = ["PDF"]
        if slides:
            bits.append(f"{len(slides)} {ru_plural(len(slides), 'слайд', 'слайда', 'слайдов')}")
        return {
            "category": category,
            "title": title,
            "subtitle": " · ".join(bits),
            "type": "pdf",
            "src": web_path(doc),
            "file": None,
            "slides": slides,
            "files_count": files_count,
        }

    if pptxs:
        doc = pptxs[0]
        slides = image_slides[:]
        try:
            slides = [web_path(p) for p in make_pptx_previews(doc, preview_root / "pptx")] + slides
        except Exception as exc:
            print(f"  ! PPTX превью ({doc.name}): {exc}")
        bits = ["PowerPoint"]
        if slides:
            bits.append(f"{len(slides)} {ru_plural(len(slides), 'слайд', 'слайда', 'слайдов')}")
            tile_type = "carousel"
        else:
            tile_type = "file"
        return {
            "category": category,
            "title": title,
            "subtitle": " · ".join(bits),
            "type": tile_type,
            "src": None,
            "file": web_path(doc),
            "slides": slides,
            "files_count": files_count,
        }

    n = len(images)
    return {
        "category": category,
        "title": title,
        "subtitle": f"{n} {ru_plural(n, 'работа', 'работы', 'работ')}",
        "type": "carousel",
        "src": None,
        "file": None,
        "slides": image_slides,
        "files_count": files_count,
    }


def scan_works() -> dict:
    WORKS.mkdir(parents=True, exist_ok=True)
    categories = []
    tiles = []

    for cat_dir in child_dirs(WORKS):
        category = cat_dir.name
        cat_tiles = 0
        for tile_dir in child_dirs(cat_dir):
            print(f"  {category}/{tile_dir.name}")
            tile = build_tile(category, tile_dir)
            if tile:
                tiles.append(tile)
                cat_tiles += 1
        if cat_tiles:
            categories.append({"id": category, "name": category})

    return {
        "categories": categories,
        "tiles": tiles,
        "stats": {
            "tiles": len(tiles),
            "categories": len(categories),
            "files": sum(t["files_count"] for t in tiles),
            "documents": sum(1 for t in tiles if t.get("src") or t.get("file")),
            "slides": sum(len(t["slides"]) for t in tiles),
        },
    }


# ── HTML render ──────────────────────────────────────────────────────────────

def render_filters(categories: list[dict]) -> str:
    buttons = ['          <button class="filter-btn active" data-filter="all">Все</button>']
    for cat in categories:
        buttons.append(
            f'          <button class="filter-btn" data-filter="{esc(cat["id"])}">{esc(cat["name"])}</button>'
        )
    return "\n".join(buttons)


def render_tile(tile: dict) -> str:
    slides = tile.get("slides") or []
    attrs = [
        f'class="gallery-item reveal"',
        f'data-category="{esc(tile["category"])}"',
        f'data-type="{esc(tile["type"])}"',
        f'data-title="{esc(tile["title"])}"',
    ]
    if tile.get("src"):
        attrs.append(f'data-src="{esc(tile["src"])}"')
    if tile.get("file"):
        attrs.append(f'data-file="{esc(tile["file"])}"')
    if slides:
        attrs.append(f'data-slides="{esc("|".join(slides))}"')

    if slides:
        imgs = "".join(
            f'<img class="slide{" active" if i == 0 else ""}" src="{esc(src)}" alt="Слайд {i + 1}">'
            for i, src in enumerate(slides)
        )
        dots = "".join(
            f'<button class="slide-dot{" active" if i == 0 else ""}" aria-label="Слайд {i + 1}"></button>'
            for i in range(len(slides))
        )
        thumb = f"""              <div class="slide-carousel">
                {imgs}
                <span class="slide-counter">1 / {len(slides)}</span>
                <div class="slide-dots">{dots}</div>
              </div>"""
    else:
        label = "PDF" if tile["type"] == "pdf" else ("PPTX" if tile.get("file") else "Файл")
        thumb = f"""              <div class="slide-carousel file-thumb">{label}</div>"""

    return f"""          <article {" ".join(attrs)}>
            <div class="gallery-thumb">
{thumb}
            </div>
            <div class="gallery-info">
              <h3>{esc(tile["title"])}</h3>
              <p>{esc(tile.get("subtitle") or tile["category"])}</p>
            </div>
          </article>"""


def render_gallery(tiles: list[dict]) -> str:
    if not tiles:
        body = """          <div class="empty-state" id="emptyState">
            <span>📂</span>
            Добавьте работы в Works/Категория/Плитка и запустите python _sync.py
          </div>"""
    else:
        cards = "\n".join(render_tile(t) for t in tiles)
        body = f"""{cards}

          <div class="empty-state hidden" id="emptyState">
            <span>📂</span>
            Работы в этой категории скоро появятся
          </div>"""
    return body


def replace_marked(text: str, name: str, content: str) -> str:
    start = f"<!-- works:{name}:start -->"
    end = f"<!-- works:{name}:end -->"
    if start not in text or end not in text:
        raise RuntimeError(f"Не найдены маркеры {start} … {end} в index.html")
    i = text.index(start) + len(start)
    j = text.index(end)
    return text[:i] + "\n" + content + "\n        " + text[j:]


def update_index(data: dict):
    text = INDEX.read_text(encoding="utf-8")
    stats = data["stats"]
    tiles = data["tiles"]

    text = replace_marked(text, "filters", render_filters(data["categories"]))
    text = replace_marked(text, "gallery", render_gallery(tiles))

    text = re.sub(
        r'(id="heroFilesCount">)\d+(</)',
        rf'\g<1>{stats["files"]}\g<2>',
        text,
        count=1,
    )
    text = re.sub(
        r'(id="heroCatsCount">)\d+(</)',
        rf'\g<1>{stats["categories"]}\g<2>',
        text,
        count=1,
    )
    for key in ("categories", "tiles", "documents", "files"):
        text = re.sub(
            rf'(data-stat="{key}">)\d+(</)',
            rf'\g<1>{stats.get(key, 0)}\g<2>',
            text,
            count=1,
        )

    INDEX.write_text(text, encoding="utf-8")


def sync() -> dict:
    print(f"Scanning {WORKS} …")
    data = scan_works()
    update_index(data)
    s = data["stats"]
    print(f"built: categories={s['categories']}, tiles={s['tiles']}, files={s['files']}")
    return data


def main():
    parser = argparse.ArgumentParser(
        description="Собрать статический сайт из папки Works/"
    )
    parser.add_argument(
        "--open",
        action="store_true",
        help="Открыть index.html после сборки",
    )
    args = parser.parse_args()
    sync()
    if args.open:
        webbrowser.open(INDEX.resolve().as_uri())


if __name__ == "__main__":
    main()
