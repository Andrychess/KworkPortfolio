import os
import re
import shutil
import fitz
from pathlib import Path

ROOT = Path(r"c:\Users\Андрей\Desktop\Портфолио")
SITE = Path(__file__).parent
WORKS = SITE / "works"

CONTACTS = {
    "name": "Пономарев Андрей",
    "short_name": "Андрей",
    "email": "andryches@gmail.com",
    "vk": "https://vk.com/andrychess",
    "vk_label": "@andrychess",
    "max": "https://max.ru/u/f9LHodD0cOK2aMHnbSYl99ctayZWMfTz1a0R5eGFbwTDvl2l9YsrGNcV2ho",
    "max_label": "Max",
}

PRESENTATIONS = {
    "Больше, чем просто работа.pdf": ("bol-she-chem-rabota", "Больше, чем просто работа", "Дизайн слайдов", "university"),
    "Конкурс команд ВУЗов по молодежной политике и воспитательной деятельности.pdf": ("konkurs-komand-vuzov", "Конкурс команд ВУЗов", "Молодёжная политика", "university"),
    "О РАБОТЕ СТРУКТУРНЫХ ПОДРАЗДЕЛЕНИЙ ВУЗА ПО ПРОФИЛАКТИКЕ ИДЕОЛОГИИ ЭКСТРЕМИЗМА И ТЕРРОРИЗМА В МОЛОДЕЖНОЙ СРЕДЕ.pdf": ("profilaktika-ekstremizma", "Профилактика экстремизма", "Структурные подразделения ВУЗа", "university"),
    "Презентация Грант СО.Знание.pdf": ("grant-so-znanie", "Грант СО.Знание", "Студенческие отряды", "university"),
    "РЕАЛИЗАЦИЯ МЕР ПОДДЕРЖКИ МОЛОДЫХ СТУДЕНЧЕСКИХ СЕМЕЙ В ФЕДЕРАЛЬНОМ ГОСУДАРСТВЕННОМ БЮДЖЕТНОМ ОБРАЗОВАТЕЛЬНОМ УЧРЕЖДЕНИИ ВЫСШЕГО ОБРАЗОВАНИЯ.pdf": ("podderzhka-student-semei", "Поддержка студенческих семей", "Меры поддержки", "university"),
    "реализация направлений деятельности в рамках полномочий проректора по молодежной и социальной полити.pdf": ("napravleniya-deyatelnosti-prorektor", "Направления деятельности", "Молодёжная и социальная политика", "university"),
    "Модный букет.pdf": ("modnyj-buket", "Модный букет", "Креативный проект", "commercial"),
    "Спика.pdf": ("spika", "Спика", "Логистика из Китая", "commercial"),
    "План развития 2024.pptx": ("plan-razvitiya-2024", "План развития 2024", "Стратегия ВУЗа", "university"),
    "Ректорская.pptx": ("rektorskaya", "Ректорская", "Официальная презентация", "university"),
}

VECTOR_RENAME = {"Знак отличия ССервО.png": "znak-servo.png"}

VECTOR_SORT = [
    "Group 695.png",
    "znak-servo.png",
    "Group 618.png",
    "Group 127.png",
    "A5 - 8.png",
    "A5 - 6.png",
    "Group 66.png",
]

FOLDER_TILES = [
    {
        "slug": "vector",
        "src": "Отрисовка в векторе",
        "dst": "vector",
        "title": "Отрисовка в векторе",
        "subtitle": "Иллюстрации, знаки отличия, стикеры",
        "category": "vector",
        "rename": VECTOR_RENAME,
        "sort": lambda f: VECTOR_SORT.index(f) if f in VECTOR_SORT else 999,
    },
    {
        "slug": "logos",
        "src": "Логотипы",
        "dst": "logos",
        "title": "Логотипы",
        "subtitle": "Фирменные знаки и эмблемы",
        "category": "branding",
    },
    {
        "slug": "icons",
        "src": "Иконки",
        "dst": "icons",
        "title": "Иконки",
        "subtitle": "Набор для презентаций и материалов",
        "category": "branding",
    },
]

SOCIAL_PROJECTS = [
    {
        "slug": "spika-1",
        "src": "Посты/Спика_1",
        "title": "Как искать товар на 1688.com",
        "subtitle": "Spika Logistics · серия 1",
        "sort": lambda f: (0, int(re.search(r"\d+", f).group()) if re.search(r"\d+", f) else f),
    },
    {
        "slug": "spika-2",
        "src": "Посты/Спика_2",
        "title": "1688.com — продолжение",
        "subtitle": "Spika Logistics · серия 2",
        "sort": lambda f: (0, int(re.search(r"\d+", f).group()) if re.search(r"\d+", f) else f),
    },
    {
        "slug": "spika-3",
        "src": "Посты/Спика_3",
        "title": "1688.com — финал серии",
        "subtitle": "Spika Logistics · серия 3",
        "sort": lambda f: (0, int(re.search(r"\d+", f).group()) if re.search(r"\d+", f) else f),
    },
    {
        "slug": "nastavnik",
        "src": "Посты/3",
        "title": "Наставник и СО.Выезд",
        "subtitle": "Посты · студотряды",
    },
    {
        "slug": "konkurs",
        "src": "Посты/4",
        "title": "Конкурс",
        "subtitle": "Серия постов · студотряды",
    },
    {
        "slug": "misc-posts",
        "src": "Посты/1",
        "title": "Разные посты",
        "subtitle": "Оформление ленты",
        "sort": lambda f: (0, int(re.search(r"\d+", Path(f).stem.replace(".", "")).group()) if re.search(r"\d+", Path(f).stem) else f),
    },
    {
        "slug": "single-post",
        "src": "Посты/2",
        "title": "Пост · коллаж",
        "subtitle": "Оформление ленты",
    },
    {
        "slug": "stories",
        "src": "Истории/И_1",
        "title": "Stories · недвижимость",
        "subtitle": "Вертикальные сторис",
        "recursive": True,
    },
]


def find_presentation(filename):
    direct = ROOT / "Презентации" / filename
    if direct.exists():
        return direct
    for path in (ROOT / "Презентации").rglob(filename):
        return path
    return None


def collect_images(folder, recursive=False, sort_key=None):
    if not folder.exists():
        return []
    files = []
    if recursive:
        for path in folder.rglob("*"):
            if path.is_file() and path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}:
                files.append(path.relative_to(folder).as_posix())
    else:
        files = [
            f for f in os.listdir(folder)
            if (folder / f).is_file() and f.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))
        ]
    if sort_key:
        files.sort(key=sort_key)
    else:
        files.sort()
    return files


def copy_folder_images(src_dir, dst_dir, rename=None, recursive=False, sort_key=None):
    rename = rename or {}
    dst_dir.mkdir(parents=True, exist_ok=True)
    copied = []
    for rel in collect_images(src_dir, recursive=recursive, sort_key=sort_key):
        src = src_dir / rel
        safe_name = rename.get(Path(rel).name, Path(rel).name)
        if rel != Path(rel).name:
            safe_name = rel.replace("/", "__")
        dst = dst_dir / safe_name
        shutil.copy2(src, dst)
        copied.append(safe_name)
    return copied


def pdf_previews(pdf_path, out_dir, count=5):
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("slide-*.png"):
        old.unlink()
    doc = fitz.open(pdf_path)
    pages = min(count, doc.page_count)
    for i in range(pages):
        page = doc[i]
        zoom = 480 / page.rect.width
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        pix.save(out_dir / f"slide-{i + 1}.png")
    doc.close()
    return pages


def pptx_previews(pptx_path, out_dir, count=5):
    import win32com.client

    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("slide-*.png"):
        old.unlink()

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 1
    presentation = app.Presentations.Open(str(pptx_path.resolve()), WithWindow=False)
    total = min(count, presentation.Slides.Count)
    for i in range(1, total + 1):
        out_path = str((out_dir / f"slide-{i}.png").resolve())
        presentation.Slides(i).Export(out_path, "PNG")
    presentation.Close()
    app.Quit()
    return total


def sync_files():
    folder_tiles = []

    for cfg in FOLDER_TILES:
        src_dir = ROOT / cfg["src"]
        dst_dir = WORKS / cfg["dst"]
        if dst_dir.exists():
            shutil.rmtree(dst_dir)
        files = copy_folder_images(
            src_dir,
            dst_dir,
            rename=cfg.get("rename"),
            sort_key=cfg.get("sort"),
        )
        if cfg.get("slug") == "vector" and files:
            files.sort(key=lambda f: VECTOR_SORT.index(f) if f in VECTOR_SORT else 999)
        if files:
            folder_tiles.append({
                "title": cfg["title"],
                "subtitle": cfg["subtitle"],
                "category": cfg["category"],
                "web_prefix": f"works/{cfg['dst']}",
                "files": files,
            })

    social_root = WORKS / "social"
    if social_root.exists():
        shutil.rmtree(social_root)

    social_tiles = []
    for project in SOCIAL_PROJECTS:
        src_dir = ROOT / "Оформление социальных сетей" / project["src"]
        dst_dir = social_root / project["slug"]
        files = copy_folder_images(
            src_dir,
            dst_dir,
            recursive=project.get("recursive", False),
            sort_key=project.get("sort"),
        )
        if files:
            social_tiles.append({
                "title": project["title"],
                "subtitle": project["subtitle"],
                "category": "social",
                "web_prefix": f"works/social/{project['slug']}",
                "files": files,
            })

    pres_dir = WORKS / "presentations"
    pres_dir.mkdir(parents=True, exist_ok=True)
    preview_root = pres_dir / "previews"

    presentations = []
    for orig, (slug, title, subtitle, group) in PRESENTATIONS.items():
        src = find_presentation(orig)
        if not src:
            continue
        ext = src.suffix.lower()
        dst = pres_dir / f"{slug}{ext}"
        shutil.copy2(src, dst)
        previews = []
        if ext == ".pdf":
            pages = pdf_previews(dst, preview_root / slug)
            previews = [f"works/presentations/previews/{slug}/slide-{i}.png" for i in range(1, pages + 1)]
        elif ext == ".pptx":
            pages = pptx_previews(dst, preview_root / slug)
            previews = [f"works/presentations/previews/{slug}/slide-{i}.png" for i in range(1, pages + 1)]
        presentations.append({
            "slug": slug,
            "title": title,
            "subtitle": subtitle,
            "group": group,
            "src": f"works/presentations/{slug}{ext}",
            "type": "pdf" if ext == ".pdf" else "pptx",
            "previews": previews,
        })

    return presentations, social_tiles, folder_tiles


def works_word(n):
    n = abs(n) % 100
    if 11 <= n <= 19:
        return "работ"
    n = n % 10
    if n == 1:
        return "работа"
    if 2 <= n <= 4:
        return "работы"
    return "работ"


def folder_carousel_card(tile):
    srcs = [f"{tile['web_prefix']}/{name}" for name in tile["files"]]
    count = len(srcs)
    subtitle = f"{tile['subtitle']} · {count} {works_word(count)}"
    return carousel_card(srcs, tile["title"], subtitle, tile["category"])


def carousel_card(srcs, title, subtitle, category):
    slides = "".join(
        f'<img class="slide{" active" if i == 0 else ""}" src="{src}" alt="Слайд {i + 1}">'
        for i, src in enumerate(srcs)
    )
    dots = "".join(
        f'<button class="slide-dot{" active" if i == 0 else ""}" aria-label="Слайд {i + 1}"></button>'
        for i in range(len(srcs))
    )
    slides_attr = "|".join(srcs)
    return f"""          <article class="gallery-item reveal" data-category="{category}" data-type="carousel"
            data-slides="{slides_attr}"
            data-title="{title}">
            <div class="gallery-thumb">
              <div class="slide-carousel">
                {slides}
                <span class="slide-counter">1 / {len(srcs)}</span>
                <div class="slide-dots">{dots}</div>
              </div>
            </div>
            <div class="gallery-info">
              <h3>{title}</h3>
              <p>{subtitle}</p>
            </div>
          </article>"""


def presentation_card(p):
    previews = p["previews"]
    if not previews:
        return ""

    slides = "".join(
        f'<img class="slide{" active" if i == 0 else ""}" src="{src}" alt="Слайд {i + 1}">'
        for i, src in enumerate(previews[:5])
    )
    dots = "".join(
        f'<button class="slide-dot{" active" if i == 0 else ""}" aria-label="Слайд {i + 1}"></button>'
        for i in range(min(5, len(previews)))
    )
    slides_attr = "|".join(previews)
    file_label = "PDF" if p["type"] == "pdf" else "PowerPoint"

    if p["type"] == "pdf":
        open_type = "pdf"
        data_open = f'\n            data-src="{p["src"]}"'
        data_slides = ""
        data_file = ""
    else:
        open_type = "carousel"
        data_open = ""
        data_slides = f'\n            data-slides="{slides_attr}"'
        data_file = f'\n            data-file="{p["src"]}"'

    return f"""          <article class="gallery-item reveal" data-category="presentations" data-type="{open_type}"{data_open}{data_slides}{data_file}
            data-title="{p['title']}">
            <div class="gallery-thumb">
              <div class="slide-carousel">
                {slides}
                <span class="slide-counter">1 / {min(5, len(previews))}</span>
                <div class="slide-dots">{dots}</div>
              </div>
            </div>
            <div class="gallery-info">
              <h3>{p['title']}</h3>
              <p>{p['subtitle']} · {file_label}</p>
            </div>
          </article>"""


def build_gallery(presentations, social_tiles, folder_tiles):
    blocks = []

    for tile in social_tiles:
        blocks.append(folder_carousel_card(tile))

    vector_tiles = [t for t in folder_tiles if t["category"] == "vector"]
    branding_tiles = [t for t in folder_tiles if t["category"] == "branding"]

    for tile in vector_tiles:
        blocks.append(folder_carousel_card(tile))

    for tile in branding_tiles:
        blocks.append(folder_carousel_card(tile))

    uni = [p for p in presentations if p["group"] == "university"]
    commercial = [p for p in presentations if p["group"] == "commercial"]
    for p in uni + commercial:
        card = presentation_card(p)
        if card:
            blocks.append(card)

    social_files = sum(len(t["files"]) for t in social_tiles)
    graphic_files = sum(len(t["files"]) for t in folder_tiles)
    total_files = social_files + graphic_files + len(presentations)
    tile_count = len(blocks)

    gallery = f"""        <div class="gallery" id="gallery">

{chr(10).join(blocks)}

          <div class="empty-state hidden" id="emptyState">
            <span>📂</span>
            Работы в этой категории скоро появятся
          </div>

        </div>"""

    index = SITE / "index.html"
    text = index.read_text(encoding="utf-8")
    start = text.index('        <div class="gallery" id="gallery">')
    end_marker = '<div class="empty-state hidden" id="emptyState">'
    end = text.index("        </div>", text.index(end_marker)) + len("        </div>")
    text = text[:start] + gallery + text[end:]

    text = re.sub(
        r"<strong>\d+</strong>\s*\n\s*работ в портфолио",
        f"<strong>{total_files}</strong>\n            работ в портфолио",
        text,
    )
    text = re.sub(
        r'<div class="stat-value">\d+</div>\s*\n\s*<div class="stat-label">Графических работ</div>',
        f'<div class="stat-value">{graphic_files}</div>\n                <div class="stat-label">Графических работ</div>',
        text,
        count=1,
    )
    text = re.sub(
        r'<div class="stat-value">\d+</div>\s*\n\s*<div class="stat-label">Презентаций</div>',
        f'<div class="stat-value">{len(presentations)}</div>\n                <div class="stat-label">Презентаций</div>',
        text,
        count=1,
    )
    text = re.sub(
        r'<div class="stat-value">\d+</div>\s*\n\s*<div class="stat-label">Работ для соцсетей</div>',
        f'<div class="stat-value">{social_files}</div>\n                <div class="stat-label">Работ для соцсетей</div>',
        text,
        count=1,
    )
    text = re.sub(
        r'<div class="stat-value">\d+</div>\s*\n\s*<div class="stat-label">Всего в портфолио</div>',
        f'<div class="stat-value">{tile_count}</div>\n                <div class="stat-label">Проектов в галерее</div>',
        text,
        count=1,
    )

    index.write_text(text, encoding="utf-8")
    print(
        f"synced: tiles={tile_count}, files={total_files}, "
        f"social={social_files}, graphic={graphic_files}, presentations={len(presentations)}"
    )


def update_contacts():
    index = SITE / "index.html"
    text = index.read_text(encoding="utf-8")
    c = CONTACTS

    text = re.sub(r"<title>.*?</title>", f"<title>{c['name']} — Графический дизайнер</title>", text, count=1)
    text = re.sub(
        r'<div class="contact-links">.*?</div>',
        f"""<div class="contact-links">
            <a href="mailto:{c['email']}" class="contact-link">✉️ {c['email']}</a>
            <a href="{c['vk']}" class="contact-link" target="_blank" rel="noopener">📘 VK · {c['vk_label']}</a>
            <a href="{c['max']}" class="contact-link" target="_blank" rel="noopener">💬 {c['max_label']}</a>
          </div>""",
        text,
        count=1,
        flags=re.DOTALL,
    )
    text = re.sub(r"© 2026 .*?\. Все права защищены\.", f"© 2026 {c['name']}. Все права защищены.", text, count=1)
    text = re.sub(
        r"Привет! Я .*? — графический дизайнер\.",
        f"Привет! Я {c['short_name']} — графический дизайнер.",
        text,
        count=1,
    )
    index.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    presentations, social_tiles, folder_tiles = sync_files()
    build_gallery(presentations, social_tiles, folder_tiles)
    update_contacts()
