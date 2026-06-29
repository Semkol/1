from __future__ import annotations

import html
import zipfile
from datetime import datetime, timezone
from pathlib import Path


OUT = Path("moon_presentation_ru.pptx")
SLIDE_W = 13_333_000
SLIDE_H = 7_500_000


def esc(text: str) -> str:
    return html.escape(text, quote=True)


def rels(items: list[tuple[str, str, str]]) -> str:
    body = "\n".join(
        f'<Relationship Id="{rid}" Type="{typ}" Target="{target}"/>'
        for rid, typ, target in items
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{body}</Relationships>"
    )


def solid_fill(color: str) -> str:
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def text_body(
    text: str,
    *,
    size: int,
    color: str = "FFFFFF",
    bold: bool = False,
    align: str = "l",
) -> str:
    bold_attr = ' b="1"' if bold else ""
    paragraphs = text.split("\n")
    p_xml: list[str] = []
    for paragraph in paragraphs:
        if not paragraph:
            p_xml.append("<a:p/>")
            continue
        p_xml.append(
            "<a:p>"
            f'<a:pPr algn="{align}"/>'
            "<a:r>"
            f'<a:rPr lang="ru-RU" sz="{size}"{bold_attr}>{solid_fill(color)}</a:rPr>'
            f"<a:t>{esc(paragraph)}</a:t>"
            "</a:r>"
            "</a:p>"
        )
    return (
        '<p:txBody><a:bodyPr wrap="square" rtlCol="0"/>'
        '<a:lstStyle/>'
        f"{''.join(p_xml)}"
        "</p:txBody>"
    )


def shape(
    sid: int,
    name: str,
    x: int,
    y: int,
    cx: int,
    cy: int,
    *,
    fill: str,
    geom: str = "rect",
    line: str | None = None,
    alpha: int | None = None,
) -> str:
    fill_xml = (
        f'<a:solidFill><a:srgbClr val="{fill}">'
        f"{f'<a:alpha val=\"{alpha}\"/>' if alpha is not None else ''}"
        "</a:srgbClr></a:solidFill>"
    )
    line_xml = f'<a:ln>{solid_fill(line)}</a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
    return (
        "<p:sp>"
        f'<p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/>'
        '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f"{fill_xml}<a:prstGeom prst=\"{geom}\"><a:avLst/></a:prstGeom>{line_xml}</p:spPr>"
        "</p:sp>"
    )


def textbox(
    sid: int,
    name: str,
    x: int,
    y: int,
    cx: int,
    cy: int,
    text: str,
    *,
    size: int,
    color: str = "FFFFFF",
    bold: bool = False,
    align: str = "l",
) -> str:
    return (
        "<p:sp>"
        f'<p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
        f"{text_body(text, size=size, color=color, bold=bold, align=align)}"
        "</p:sp>"
    )


def bullet_list(
    sid_start: int,
    x: int,
    y: int,
    cx: int,
    bullets: list[str],
    *,
    color: str = "E9EEF8",
) -> str:
    shapes: list[str] = []
    gap = 590_000
    for i, bullet in enumerate(bullets):
        yy = y + i * gap
        shapes.append(shape(sid_start + i * 2, "bullet", x, yy + 100_000, 95_000, 95_000, fill="F7C948", geom="ellipse"))
        shapes.append(
            textbox(
                sid_start + i * 2 + 1,
                "bullet text",
                x + 210_000,
                yy,
                cx - 210_000,
                410_000,
                bullet,
                size=2100,
                color=color,
            )
        )
    return "".join(shapes)


def title_block(title: str, subtitle: str | None = None) -> str:
    parts = [
        textbox(20, "title", 670_000, 460_000, 8_900_000, 700_000, title, size=3900, bold=True)
    ]
    if subtitle:
        parts.append(textbox(21, "subtitle", 690_000, 1_220_000, 8_400_000, 430_000, subtitle, size=2000, color="BFD7FF"))
    return "".join(parts)


def moon_art(x: int, y: int, size: int, sid: int = 80) -> str:
    craters = [
        (0.20, 0.19, 0.13),
        (0.54, 0.16, 0.08),
        (0.66, 0.42, 0.15),
        (0.31, 0.55, 0.10),
        (0.50, 0.70, 0.12),
        (0.18, 0.76, 0.06),
    ]
    parts = [
        shape(sid, "Moon", x, y, size, size, fill="D9DEE8", geom="ellipse", line="F8FAFC"),
        shape(sid + 1, "Moon shadow", x + int(size * 0.58), y + int(size * 0.04), int(size * 0.37), int(size * 0.92), fill="AAB2C2", geom="ellipse", alpha=25000),
    ]
    for idx, (cx, cy, r) in enumerate(craters, start=2):
        d = int(size * r)
        parts.append(
            shape(
                sid + idx,
                "Crater",
                x + int(size * cx),
                y + int(size * cy),
                d,
                d,
                fill="98A2B3",
                geom="ellipse",
                alpha=55000,
            )
        )
    return "".join(parts)


def phase_moons() -> str:
    names = ["новолуние", "первая\nчетверть", "полнолуние", "последняя\nчетверть"]
    fills = ["1F2937", "B8C0CC", "E5E7EB", "8791A3"]
    parts: list[str] = []
    for i, (name, fill) in enumerate(zip(names, fills)):
        x = 1_360_000 + i * 2_550_000
        parts.append(shape(140 + i, "phase", x, 4_950_000, 770_000, 770_000, fill=fill, geom="ellipse", line="E5E7EB"))
        parts.append(textbox(150 + i, "phase label", x - 190_000, 5_820_000, 1_160_000, 450_000, name, size=1500, color="D7E3F7", align="ctr"))
    return "".join(parts)


def slide_xml(title: str, subtitle: str | None, bullets: list[str], art: str = "") -> str:
    shapes = [
        shape(2, "background", 0, 0, SLIDE_W, SLIDE_H, fill="0B1020"),
        shape(3, "accent", 0, 0, 260_000, SLIDE_H, fill="F7C948"),
        shape(4, "top glow", 6_400_000, -1_000_000, 4_500_000, 4_500_000, fill="1E3A8A", geom="ellipse", alpha=35000),
        title_block(title, subtitle),
        bullet_list(30, 760_000, 2_080_000, 7_500_000, bullets),
        art or moon_art(9_360_000, 1_850_000, 2_900_000),
    ]
    return wrap_slide("".join(shapes))


def wrap_slide(sp_tree: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f"{sp_tree}"
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    )


def content_types(slide_count: int) -> str:
    overrides = [
        ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        ("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
        ("/ppt/slideMasters/slideMaster1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"),
        ("/ppt/slideLayouts/slideLayout1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"),
        ("/ppt/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
        ("/ppt/presProps.xml", "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"),
        ("/ppt/viewProps.xml", "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"),
        ("/ppt/tableStyles.xml", "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"),
    ]
    overrides.extend(
        (f"/ppt/slides/slide{i}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
        for i in range(1, slide_count + 1)
    )
    override_xml = "".join(f'<Override PartName="{part}" ContentType="{ctype}"/>' for part, ctype in overrides)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{override_xml}</Types>"
    )


PRESENTATION_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
    '<p:sldIdLst>{slide_ids}</p:sldIdLst>'
    f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="wide"/>'
    '<p:notesSz cx="6858000" cy="9144000"/>'
    '<p:defaultTextStyle><a:defPPr><a:defRPr lang="ru-RU"/></a:defPPr></p:defaultTextStyle>'
    '</p:presentation>'
)


SLIDE_LAYOUT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">'
    '<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
    '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
    '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>'
    '</p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
)


SLIDE_MASTER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
    '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
    '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>'
    '</p:grpSpPr></p:spTree></p:cSld>'
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
    'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>'
)


THEME = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Moon Theme">'
    '<a:themeElements><a:clrScheme name="Moon">'
    '<a:dk1><a:srgbClr val="0B1020"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="E5E7EB"/></a:lt2>'
    '<a:accent1><a:srgbClr val="F7C948"/></a:accent1><a:accent2><a:srgbClr val="60A5FA"/></a:accent2>'
    '<a:accent3><a:srgbClr val="94A3B8"/></a:accent3><a:accent4><a:srgbClr val="A78BFA"/></a:accent4>'
    '<a:accent5><a:srgbClr val="22D3EE"/></a:accent5><a:accent6><a:srgbClr val="F472B6"/></a:accent6>'
    '<a:hlink><a:srgbClr val="60A5FA"/></a:hlink><a:folHlink><a:srgbClr val="A78BFA"/></a:folHlink>'
    '</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
    '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
)


def build() -> None:
    slides = [
        wrap_slide(
            "".join(
                [
                    shape(2, "background", 0, 0, SLIDE_W, SLIDE_H, fill="070B18"),
                    shape(3, "gold line", 0, 6_900_000, SLIDE_W, 90_000, fill="F7C948"),
                    textbox(20, "title", 780_000, 970_000, 6_900_000, 900_000, "Луна", size=5600, bold=True),
                    textbox(21, "subtitle", 820_000, 1_870_000, 6_800_000, 520_000, "Ближайший естественный спутник Земли", size=2200, color="BFD7FF"),
                    textbox(22, "caption", 820_000, 5_850_000, 6_300_000, 520_000, "Краткая презентация о происхождении, свойствах и значении Луны", size=1800, color="D7E3F7"),
                    moon_art(8_100_000, 860_000, 4_230_000),
                    shape(90, "earth", 10_920_000, 5_500_000, 930_000, 930_000, fill="2563EB", geom="ellipse", line="93C5FD"),
                    shape(91, "earth land", 11_150_000, 5_670_000, 310_000, 210_000, fill="22C55E", geom="ellipse", alpha=70000),
                ]
            )
        ),
        slide_xml(
            "Что такое Луна",
            "Основные характеристики",
            [
                "Среднее расстояние до Земли — около 384 400 км.",
                "Диаметр — 3 474 км, примерно четверть диаметра Земли.",
                "Сила тяжести на поверхности почти в 6 раз меньше земной.",
                "Луна всегда обращена к Земле почти одной стороной.",
            ],
        ),
        slide_xml(
            "Как Луна появилась",
            "Гипотеза гигантского столкновения",
            [
                "Около 4,5 млрд лет назад молодая Земля столкнулась с крупным телом.",
                "Выброшенное вещество образовало диск вокруг Земли.",
                "Из этого диска постепенно сформировалась Луна.",
                "Состав лунных пород помогает проверять эту гипотезу.",
            ],
        ),
        slide_xml(
            "Поверхность Луны",
            "Моря, кратеры и реголит",
            [
                "Тёмные «моря» — древние лавовые равнины, а не вода.",
                "Кратеры образованы ударами метеоритов и астероидов.",
                "Поверхность покрыта реголитом — слоем пыли и обломков.",
                "Из-за слабой атмосферы перепады температур очень велики.",
            ],
        ),
        slide_xml(
            "Фазы Луны",
            "Почему её вид меняется",
            [
                "Луна светит отражённым солнечным светом.",
                "Фаза зависит от взаимного положения Солнца, Земли и Луны.",
                "Полный цикл фаз длится примерно 29,5 суток.",
                "Затмения происходят, когда тела выстраиваются почти в одну линию.",
            ],
            art=phase_moons(),
        ),
        slide_xml(
            "Влияние на Землю",
            "Луна заметно меняет нашу планету",
            [
                "Гравитация Луны вызывает океанские приливы и отливы.",
                "Она помогает стабилизировать наклон земной оси.",
                "Лунный цикл лежит в основе многих календарей.",
                "Ночная освещённость Луны важна для природы и культуры.",
            ],
        ),
        slide_xml(
            "Исследование Луны",
            "От первых аппаратов до новых миссий",
            [
                "Советские станции «Луна» первыми достигли поверхности и доставили образцы.",
                "В 1969 году Apollo 11 впервые высадил людей на Луну.",
                "Орбитальные аппараты составили подробные карты поверхности.",
                "Современные программы изучают ресурсы и возможность лунных баз.",
            ],
        ),
        slide_xml(
            "Почему Луна важна",
            "Наука, технологии и будущее",
            [
                "Лунные породы хранят историю ранней Солнечной системы.",
                "Полярный лёд может стать ресурсом для будущих экспедиций.",
                "Луна удобна для испытания технологий дальнего космоса.",
                "Её изучение помогает лучше понять Землю.",
            ],
        ),
    ]

    slide_ids = "".join(f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, len(slides) + 1))
    presentation_rels = [("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "slideMasters/slideMaster1.xml")]
    presentation_rels.extend(
        (
            f"rId{i + 1}",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
            f"slides/slide{i}.xml",
        )
        for i in range(1, len(slides) + 1)
    )
    presentation_rels.extend(
        [
            (f"rId{len(slides) + 2}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"),
            (f"rId{len(slides) + 3}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps", "presProps.xml"),
            (f"rId{len(slides) + 4}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps", "viewProps.xml"),
            (f"rId{len(slides) + 5}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles", "tableStyles.xml"),
        ]
    )

    now = datetime.now(timezone.utc).isoformat()
    with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types(len(slides)))
        zf.writestr("_rels/.rels", rels([
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "ppt/presentation.xml"),
            ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"),
            ("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml"),
        ]))
        zf.writestr(
            "docProps/core.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>Луна</dc:title><dc:subject>Презентация о Луне</dc:subject>'
            '<dc:creator>Cursor Cloud Agent</dc:creator><cp:keywords>Луна, космос, презентация</cp:keywords>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
            "</cp:coreProperties>",
        )
        zf.writestr(
            "docProps/app.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Python ZIP OOXML</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat>'
            f"<Slides>{len(slides)}</Slides><Company></Company></Properties>",
        )
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML.format(slide_ids=slide_ids))
        zf.writestr("ppt/_rels/presentation.xml.rels", rels(presentation_rels))
        zf.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", rels([
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "../theme/theme1.xml"),
        ]))
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", rels([
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "../slideMasters/slideMaster1.xml")
        ]))
        zf.writestr("ppt/theme/theme1.xml", THEME)
        zf.writestr("ppt/presProps.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')
        zf.writestr("ppt/viewProps.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')
        zf.writestr("ppt/tableStyles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>')
        for i, slide in enumerate(slides, start=1):
            zf.writestr(f"ppt/slides/slide{i}.xml", slide)
            zf.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", rels([
                ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml")
            ]))


if __name__ == "__main__":
    build()
    print(f"Created {OUT}")
