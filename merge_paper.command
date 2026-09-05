#!/bin/bash

## 需要修改：输入的两个 docx 的文件名
WORD1_NAME="word1-preface.docx"
WORD2_NAME="word2-content.docx"

## 可以修改：指定最终版文件名
OUTPUT_NAME="Dissertation-Final.pdf"

## 可以修改：书签层级，默认为三级
MAX_OUTLINE_LEVEL=3

DIR_PATH="$(cd "$(dirname "$0")" && pwd)"

WORD1="$DIR_PATH/$WORD1_NAME"
WORD2="$DIR_PATH/$WORD2_NAME"

MERGE_DIR="$DIR_PATH/merge"

COPY1="$MERGE_DIR/$WORD1_NAME"
COPY2="$MERGE_DIR/$WORD2_NAME"

PDF1="$MERGE_DIR/${WORD1_NAME%.docx}.pdf"
PDF2="$MERGE_DIR/${WORD2_NAME%.docx}.pdf"

OUTPUT="$MERGE_DIR/$OUTPUT_NAME"

PY_SCRIPT="$MERGE_DIR/create_bookmarks.py"

BOOKMARK_LOG="$OUTPUT.bookmark_log.txt"


# ============================================================
# 开始
# ============================================================

clear

echo ""
echo "=========================================="
echo " 博士论文高清书签版 PDF"
echo "=========================================="
echo ""

echo "PDF 设置："
echo "  ✓ Microsoft Word 原生导出"
echo "  ✓ 打印质量"
echo "  ✓ 图片保持高清"
echo "  ✓ 自动读取 Word 大纲级别"
echo "  ✓ 支持自定义 Word 样式"
echo "  ✓ 自动识别标题编号"
echo "  ✓ 自动排除目录中的同名标题"
echo "  ✓ 自动生成 PDF 层级书签"
echo "  ✓ 自动处理前置页页码偏移"
echo ""

echo "⚠️ 原始 Word 文件不会被打开或修改。"
echo ""


# ============================================================
# 检查 Microsoft Word
# ============================================================

if [ ! -d "/Applications/Microsoft Word.app" ]; then

    echo "❌ 找不到 Microsoft Word"
    echo ""

    read -n 1
    exit 1

fi


# ============================================================
# 检查原始 Word 文件
# ============================================================

if [ ! -f "$WORD1" ]; then

    echo "❌ 找不到："
    echo "$WORD1"
    echo ""

    read -n 1
    exit 1

fi


if [ ! -f "$WORD2" ]; then

    echo "❌ 找不到："
    echo "$WORD2"
    echo ""

    read -n 1
    exit 1

fi


# ============================================================
# 检查 Python
# ============================================================

if ! command -v python3 >/dev/null 2>&1; then

    echo "❌ 找不到 python3"
    echo ""

    read -n 1
    exit 1

fi


# ============================================================
# 检查 PyMuPDF
# ============================================================

if ! python3 -c "import fitz" >/dev/null 2>&1; then

    echo "❌ 没有安装 PyMuPDF"
    echo ""
    echo "请先在终端运行："
    echo ""
    echo "python3 -m pip install pymupdf"
    echo ""

    read -n 1
    exit 1

fi


# ============================================================
# 创建 merge 工作目录
# ============================================================

echo "① 准备 merge 工作目录..."

mkdir -p "$MERGE_DIR"

if [ $? -ne 0 ]; then

    echo ""
    echo "❌ 无法创建 merge 文件夹"
    echo ""

    read -n 1
    exit 1

fi


# ============================================================
# 清理上一次生成的工作文件
#
# 只删除 merge 中的文件
# 不碰原始 Word
# ============================================================

rm -f "$COPY1"
rm -f "$COPY2"

rm -f "$PDF1"
rm -f "$PDF2"

rm -f "$OUTPUT"

rm -f "$PY_SCRIPT"
rm -f "$BOOKMARK_LOG"


# ============================================================
# 复制 Word
# ============================================================

echo ""
echo "② 复制 Word 文件到 merge..."


cp "$WORD1" "$COPY1"

if [ $? -ne 0 ]; then

    echo ""
    echo "❌ 第一个 Word 复制失败"

    read -n 1
    exit 1

fi


cp "$WORD2" "$COPY2"

if [ $? -ne 0 ]; then

    echo ""
    echo "❌ 第二个 Word 复制失败"

    read -n 1
    exit 1

fi


echo "✅ Word 副本创建完成"


# ============================================================
# Microsoft Word → PDF
#
# 只打开 merge 中的副本
#
# 这里使用 Word 自身 PDF 引擎。
# 不经过 LibreOffice。
#
# ============================================================

export_word_pdf() {

    DOCX_PATH="$1"
    PDF_PATH="$2"
    DOC_NAME="$3"

    echo ""
    echo "Microsoft Word 正在导出："
    echo "$DOC_NAME"
    echo ""

    osascript <<EOF

tell application "Microsoft Word"

    activate

    open POSIX file "$DOCX_PATH"

    delay 3

    set theDoc to active document

    save as theDoc file name "$PDF_PATH" file format format PDF

    delay 3

    close theDoc saving no

end tell

EOF

    STATUS=$?


    if [ "$STATUS" -ne 0 ]; then

        echo ""
        echo "❌ Microsoft Word 导出失败："
        echo "$DOC_NAME"

        return 1

    fi


    if [ ! -s "$PDF_PATH" ]; then

        echo ""
        echo "❌ PDF 没有正常生成："
        echo "$PDF_PATH"

        return 1

    fi


    echo "✅ 已生成："
    echo "$PDF_PATH"

    return 0
}


# ============================================================
# 导出第一个 Word
# ============================================================

echo ""
echo "③ 导出题名页、授权书 PDF..."


export_word_pdf \
    "$COPY1" \
    "$PDF1" \
    "$WORD1_NAME"


if [ $? -ne 0 ]; then

    echo ""
    read -n 1
    exit 1

fi


# ============================================================
# 导出论文正文
# ============================================================

echo ""
echo "④ 导出论文正文 PDF..."


export_word_pdf \
    "$COPY2" \
    "$PDF2" \
    "$WORD2_NAME"


if [ $? -ne 0 ]; then

    echo ""
    read -n 1
    exit 1

fi


# ============================================================
# 创建 Python 书签程序
# ============================================================

echo ""
echo "⑤ 正在创建书签处理程序..."


cat > "$PY_SCRIPT" <<'PYTHON'
import sys
import re
import zipfile
import unicodedata
import xml.etree.ElementTree as ET

import fitz


# ============================================================
# 参数
# ============================================================

docx_path = sys.argv[1]
front_pdf_path = sys.argv[2]
body_pdf_path = sys.argv[3]
output_pdf_path = sys.argv[4]
max_level = int(sys.argv[5])


# ============================================================
# Word XML namespace
# ============================================================

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

NS = {
    "w": W
}


def wattr(name):

    return f"{{{W}}}{name}"


# ============================================================
# 文本标准化
#
# 用于 Word XML 与 PDF 文字之间的匹配。
#
# ============================================================

def normalize_text(text):

    if text is None:
        return ""

    text = unicodedata.normalize(
        "NFKC",
        text
    )

    text = text.replace(
        "\u00a0",
        " "
    )

    text = text.replace(
        "\u3000",
        " "
    )

    # 删除所有空白
    text = re.sub(
        r"\s+",
        "",
        text
    )

    return text.strip()


# ============================================================
# 读取 Word styles.xml
#
# 建立：
#
# styleId
#     ↓
# style name
# basedOn
# outline level
#
# outlineLvl：
#
# 0 → 大纲级别 1
# 1 → 大纲级别 2
# 2 → 大纲级别 3
#
# ============================================================

def read_styles(zip_file):

    styles_xml = zip_file.read(
        "word/styles.xml"
    )

    root = ET.fromstring(
        styles_xml
    )

    styles = {}

    for style in root.findall(
        "w:style",
        NS
    ):

        style_type = style.get(
            wattr("type")
        )

        if style_type != "paragraph":
            continue


        style_id = style.get(
            wattr("styleId")
        )

        if not style_id:
            continue


        # ----------------------------------------------------
        # 样式名称
        # ----------------------------------------------------

        name_node = style.find(
            "w:name",
            NS
        )

        style_name = (
            name_node.get(
                wattr("val")
            )
            if name_node is not None
            else style_id
        )


        # ----------------------------------------------------
        # basedOn
        # ----------------------------------------------------

        based_node = style.find(
            "w:basedOn",
            NS
        )

        based_on = (
            based_node.get(
                wattr("val")
            )
            if based_node is not None
            else None
        )


        # ----------------------------------------------------
        # outlineLvl
        # ----------------------------------------------------

        outline_node = style.find(
            "w:pPr/w:outlineLvl",
            NS
        )

        outline = None

        if outline_node is not None:

            value = outline_node.get(
                wattr("val")
            )

            if value is not None:

                outline = (
                    int(value)
                    + 1
                )


        styles[style_id] = {

            "name": style_name,

            "based_on": based_on,

            "outline": outline,
        }


    return styles


# ============================================================
# 解析自定义样式的大纲级别
#
# 支持 basedOn 继承。
# ============================================================

def resolve_style_outline(
    style_id,
    styles
):

    visited = set()

    current = style_id


    while (
        current
        and current not in visited
    ):

        visited.add(
            current
        )

        style = styles.get(
            current
        )

        if not style:

            return None


        if style["outline"] is not None:

            return style["outline"]


        current = style["based_on"]


    return None


# ============================================================
# 段落直接设置的大纲级别
# ============================================================

def get_direct_outline(
    paragraph
):

    node = paragraph.find(
        "w:pPr/w:outlineLvl",
        NS
    )

    if node is None:

        return None


    value = node.get(
        wattr("val")
    )

    if value is None:

        return None


    return int(value) + 1


# ============================================================
# 获取段落样式
# ============================================================

def get_paragraph_style(
    paragraph
):

    node = paragraph.find(
        "w:pPr/w:pStyle",
        NS
    )

    if node is None:

        return None


    return node.get(
        wattr("val")
    )


# ============================================================
# 获取 Word 段落文字
#
# 注意：
#
# Word 自动多级编号通常不会出现在普通 w:t 中。
#
# 因此这里得到：
#
#   绪论
#
# 而 PDF 中可能是：
#
#   第一章 绪论
#
# 后面会从 PDF 中读取完整显示标题。
#
# ============================================================

def paragraph_text(
    paragraph
):

    parts = []


    for node in paragraph.iter():

        if node.tag == f"{{{W}}}t":

            if node.text:

                parts.append(
                    node.text
                )


        elif node.tag in {

            f"{{{W}}}tab",

            f"{{{W}}}br",

        }:

            parts.append(" ")


    return "".join(
        parts
    ).strip()


# ============================================================
# 读取 DOCX 大纲
# ============================================================

def read_docx_outline(
    docx_path,
    max_level
):

    headings = []


    with zipfile.ZipFile(
        docx_path
    ) as z:


        styles = read_styles(
            z
        )


        document_xml = z.read(
            "word/document.xml"
        )


        root = ET.fromstring(
            document_xml
        )


        paragraphs = root.findall(
            ".//w:body/w:p",
            NS
        )


        for p in paragraphs:


            text = paragraph_text(
                p
            )


            if not text:
                continue


            # ------------------------------------------------
            # 优先读取段落直接大纲级别
            # ------------------------------------------------

            level = get_direct_outline(
                p
            )


            style_id = get_paragraph_style(
                p
            )


            # ------------------------------------------------
            # 如果段落本身没有，则读取样式
            # ------------------------------------------------

            if (
                level is None
                and style_id
            ):

                level = resolve_style_outline(
                    style_id,
                    styles
                )


            if level is None:
                continue


            if not (
                1
                <= level
                <= max_level
            ):

                continue


            style_name = ""


            if style_id in styles:

                style_name = (
                    styles[
                        style_id
                    ]["name"]
                )


            headings.append({

                "level": level,

                "text": text,

                "style": style_name,

            })


    return headings


# ============================================================
# 读取 PDF 每一行
#
# 同时记录：
#
#   页面
#   文字
#   坐标
#   字体大小
#
# ============================================================

def read_pdf_lines(pdf):

    pages = []


    for page_index in range(
        len(pdf)
    ):


        page = pdf[
            page_index
        ]


        data = page.get_text(
            "dict",
            sort=True
        )


        lines = []


        for block in data.get(
            "blocks",
            []
        ):


            if "lines" not in block:
                continue


            for line in block[
                "lines"
            ]:


                spans = line.get(
                    "spans",
                    []
                )


                text = "".join(

                    span.get(
                        "text",
                        ""
                    )

                    for span in spans

                ).strip()


                if not text:
                    continue


                bbox = line[
                    "bbox"
                ]


                font_sizes = [

                    span.get(
                        "size",
                        0
                    )

                    for span in spans

                    if span.get(
                        "text",
                        ""
                    ).strip()

                ]


                max_font_size = (

                    max(font_sizes)

                    if font_sizes

                    else 0

                )


                lines.append({

                    "text": text,

                    "normalized":
                        normalize_text(
                            text
                        ),

                    "x0": bbox[0],

                    "y0": bbox[1],

                    "x1": bbox[2],

                    "y1": bbox[3],

                    "font_size":
                        max_font_size,

                })


        pages.append(
            lines
        )


    return pages


# ============================================================
# 判断是不是目录中的标题
#
# 典型目录：
#
# 第一章 绪论................1
#
# 第一章 绪论………………1
#
# 第一章 绪论              1
#
# ============================================================

def is_toc_line(
    line_text,
    heading_text
):

    text = unicodedata.normalize(
        "NFKC",
        line_text
    ).strip()


    # --------------------------------------------------------
    # 目录点线
    # --------------------------------------------------------

    if re.search(
        r"[\.．…·]{3,}",
        text
    ):

        return True


    heading_norm = normalize_text(
        heading_text
    )

    line_norm = normalize_text(
        text
    )


    position = line_norm.find(
        heading_norm
    )


    if position >= 0:


        rest = line_norm[

            position
            + len(heading_norm):

        ]


        # ----------------------------------------------------
        # 标题后只剩页码
        #
        # 例如：
        #
        # 绪论1
        # 绪论12
        #
        # ----------------------------------------------------

        if (
            rest
            and re.fullmatch(
                r"[0-9０-９ivxlcdmIVXLCDM\-–—]+",
                rest
            )
        ):

            return True


    return False


# ============================================================
# 标题候选评分
#
# 目的：
#
# 目录：
#   降低优先级
#
# 页眉：
#   字体通常较小
#
# 正文标题：
#   字体通常更大
#
# ============================================================

def candidate_score(
    candidate
):

    line = candidate["line"]


    score = 0


    # --------------------------------------------------------
    # 目录直接强烈降权
    # --------------------------------------------------------

    if candidate["toc_like"]:

        score -= 10000


    # --------------------------------------------------------
    # 字体越大越优先
    # --------------------------------------------------------

    score += (
        line.get(
            "font_size",
            0
        )
        * 100
    )


    # --------------------------------------------------------
    # 页面顶部极靠上的文字
    # 很可能是页眉
    # --------------------------------------------------------

    if line.get(
        "y0",
        0
    ) < 60:

        score -= 1000


    return score


# ============================================================
# 查找一个标题
#
# 不再：
#
#   找到第一个 → 立即返回
#
# 而是：
#
#   找到所有候选
#       ↓
#   排除目录
#       ↓
#   避免页眉
#       ↓
#   优先正文标题
#
# ============================================================

def find_heading_line(
    heading_text,
    pdf_pages,
    start_page
):

    target = normalize_text(
        heading_text
    )


    if not target:

        return None


    candidates = []


    # ========================================================
    # 从当前进度向后搜索
    # ========================================================

    for page_index in range(

        start_page,

        len(pdf_pages)

    ):


        for line in pdf_pages[
            page_index
        ]:


            candidate_text = (
                line["normalized"]
            )


            if target not in candidate_text:

                continue


            toc_like = is_toc_line(

                line["text"],

                heading_text

            )


            candidates.append({

                "page":
                    page_index,

                "line":
                    line,

                "toc_like":
                    toc_like,

            })


    if not candidates:

        return None


    # ========================================================
    # 首先只保留非目录候选
    # ========================================================

    normal_candidates = [

        c

        for c in candidates

        if not c["toc_like"]

    ]


    # ========================================================
    # 有正文候选
    # ========================================================

    if normal_candidates:


        # ----------------------------------------------------
        # 评分：
        #
        # 1. 字体大小
        # 2. 避免页眉
        #
        # 同分时：
        # 页面靠前者优先
        #
        # ----------------------------------------------------

        normal_candidates.sort(

            key=lambda c: (

                -candidate_score(c),

                c["page"],

                c["line"].get(
                    "y0",
                    0
                ),

            )

        )


        return normal_candidates[0]


    # ========================================================
    # 实在没有非目录候选
    #
    # 才退回目录候选
    # ========================================================

    candidates.sort(

        key=lambda c: (

            -candidate_score(c),

            c["page"]

        )

    )


    return candidates[0]


# ============================================================
# 定位所有标题
# ============================================================

def locate_headings(
    headings,
    pdf_pages
):

    results = []

    current_page = 0


    for heading in headings:


        match = find_heading_line(

            heading["text"],

            pdf_pages,

            current_page

        )


        item = heading.copy()


        if match is None:


            item["page"] = None

            item["pdf_title"] = (
                heading["text"]
            )

            item["y"] = 0

            item["font_size"] = 0


        else:


            item["page"] = (
                match["page"]
            )


            # ------------------------------------------------
            # 使用 PDF 中 Word 实际显示的完整标题
            #
            # Word XML：
            #
            #   绪论
            #
            # PDF：
            #
            #   第一章 绪论
            #
            # ------------------------------------------------

            item["pdf_title"] = (
                match["line"]["text"]
            )


            item["y"] = max(

                match[
                    "line"
                ]["y0"] - 10,

                0

            )


            item["font_size"] = (
                match[
                    "line"
                ].get(
                    "font_size",
                    0
                )
            )


            current_page = (
                match["page"]
            )


        results.append(
            item
        )


    return results


# ============================================================
# 开始分析
# ============================================================

print()

print(
    "=========================================="
)

print(
    " Word 大纲分析"
)

print(
    "=========================================="
)

print()


# ============================================================
# Word 大纲
# ============================================================

headings = read_docx_outline(

    docx_path,

    max_level

)


if not headings:

    print(
        "❌ 没有发现 Word 大纲标题。"
    )

    print()

    print(
        "请确认自定义样式已经设置："
    )

    print(
        "格式 → 段落 → 大纲级别"
    )

    sys.exit(2)


print(
    f"✓ Word 大纲标题："
    f"{len(headings)} 个"
)


# ============================================================
# 打开正文 PDF
# ============================================================

body_pdf = fitz.open(
    body_pdf_path
)


print(
    f"✓ 正文 PDF："
    f"{len(body_pdf)} 页"
)


# ============================================================
# 读取 PDF 文字
# ============================================================

pdf_pages = read_pdf_lines(
    body_pdf
)


# ============================================================
# 定位
# ============================================================

headings = locate_headings(

    headings,

    pdf_pages

)


found = [

    h

    for h in headings

    if h["page"] is not None

]


missing = [

    h

    for h in headings

    if h["page"] is None

]


print(
    f"✓ 成功定位："
    f"{len(found)} 个"
)


if missing:

    print(
        f"⚠️ 未定位："
        f"{len(missing)} 个"
    )


# ============================================================
# 书签预览
# ============================================================

print()

print(
    "------------------------------------------"
)

print(
    "书签识别结果"
)

print(
    "------------------------------------------"
)

print()


for h in headings:


    indent = (
        "    "
        * (
            h["level"]
            - 1
        )
    )


    if h["page"] is None:


        print(

            f"{indent}"
            f"[L{h['level']}] "
            f"{h['text']} "
            f"→ 未找到"

        )


    else:


        print(

            f"{indent}"
            f"[L{h['level']}] "
            f"{h['pdf_title']} "
            f"→ 正文第 "
            f"{h['page'] + 1} 页 "
            f"(字体 "
            f"{h['font_size']:.1f})"

        )


# ============================================================
# 写日志
# ============================================================

log_path = (
    output_pdf_path
    + ".bookmark_log.txt"
)


with open(

    log_path,

    "w",

    encoding="utf-8"

) as f:


    f.write(
        "Word 大纲 → PDF 书签识别结果\n"
    )

    f.write(
        "=" * 80
        + "\n\n"
    )


    for h in headings:


        if h["page"] is None:


            f.write(

                f"L{h['level']}\t"
                f"未找到\t"
                f"{h['style']}\t"
                f"{h['text']}\n"

            )


        else:


            f.write(

                f"L{h['level']}\t"
                f"{h['page'] + 1}\t"
                f"{h['font_size']:.1f}\t"
                f"{h['style']}\t"
                f"{h['pdf_title']}\n"

            )


# ============================================================
# 合并 PDF
# ============================================================

print()

print(
    "------------------------------------------"
)

print(
    "合并 PDF"
)

print(
    "------------------------------------------"
)

print()


front_pdf = fitz.open(
    front_pdf_path
)


front_pages = len(
    front_pdf
)


print(
    f"✓ 题名页、授权书："
    f"{front_pages} 页"
)


print(
    f"✓ 论文正文："
    f"{len(body_pdf)} 页"
)


# ============================================================
# 创建最终 PDF
# ============================================================

output = fitz.open()


output.insert_pdf(
    front_pdf
)


output.insert_pdf(
    body_pdf
)


# ============================================================
# 创建 PDF TOC
#
# 使用：
#
# LINK_GOTO
#
# 明确建立内部页面跳转。
#
# ============================================================

toc = []

previous_level = 0


for h in headings:


    if h["page"] is None:

        continue


    level = h["level"]


    # --------------------------------------------------------
    # PyMuPDF 要求层级连续
    # --------------------------------------------------------

    if previous_level == 0:

        level = 1


    elif (
        level
        > previous_level + 1
    ):

        level = (
            previous_level + 1
        )


    # --------------------------------------------------------
    # 最终 PDF：
    #
    # 前置 PDF
    # +
    # 正文 PDF
    #
    # --------------------------------------------------------

    target_page_zero_based = (

        front_pages

        + h["page"]

    )


    target_page_one_based = (

        target_page_zero_based

        + 1

    )


    # --------------------------------------------------------
    # 明确创建内部跳转
    # --------------------------------------------------------

    destination = {

        "kind":
            fitz.LINK_GOTO,

        "page":
            target_page_zero_based,

        "to":
            fitz.Point(
                0,
                h["y"]
            ),

        "zoom":
            0,

    }


    toc.append([

        level,

        h["pdf_title"],

        target_page_one_based,

        destination,

    ])


    previous_level = level


# ============================================================
# 检查 TOC
# ============================================================

if not toc:

    print()

    print(
        "❌ 没有可以写入 PDF 的书签。"
    )

    front_pdf.close()
    body_pdf.close()
    output.close()

    sys.exit(3)


# ============================================================
# 写入 TOC
# ============================================================

output.set_toc(

    toc,

    collapse=0

)


# ============================================================
# 保存
#
# 不重新渲染 PDF 页面。
# 不重新采样图片。
#
# ============================================================

output.save(

    output_pdf_path,

    garbage=0,

    deflate=False,

    clean=False

)


# ============================================================
# 关闭
# ============================================================

front_pdf.close()

body_pdf.close()

output.close()


# ============================================================
# 最终验证
# ============================================================

check = fitz.open(
    output_pdf_path
)


final_toc = check.get_toc(
    simple=False
)


final_pages = len(
    check
)


valid_links = 0


print()

print(
    "------------------------------------------"
)

print(
    "最终 PDF 书签验证"
)

print(
    "------------------------------------------"
)

print()


for item in final_toc:


    level = item[0]

    title = item[1]

    page = item[2]

    dest = item[3]


    kind = dest.get(
        "kind",
        fitz.LINK_NONE
    )


    indent = (
        "    "
        * (
            level
            - 1
        )
    )


    if kind == fitz.LINK_GOTO:

        valid_links += 1

        status = "✓"


    else:

        status = "❌"


    print(

        f"{indent}"
        f"{status} "
        f"{title} "
        f"→ 最终PDF第 {page} 页"

    )


check.close()


# ============================================================
# 汇总
# ============================================================

print()

print(
    "=========================================="
)

print(
    " PDF 处理完成"
)

print(
    "=========================================="
)

print()


print(
    f"✓ 最终 PDF："
    f"{final_pages} 页"
)


print(
    f"✓ Word 大纲："
    f"{len(headings)} 个"
)


print(
    f"✓ 成功定位："
    f"{len(found)} 个"
)


print(
    f"✓ PDF 书签："
    f"{len(final_toc)} 个"
)


print(
    f"✓ 有效跳转："
    f"{valid_links} 个"
)


print(
    f"⚠️ 未定位："
    f"{len(missing)} 个"
)


print()

print(
    "书签日志："
)

print(
    log_path
)


# ============================================================
# 检查跳转
# ============================================================

if valid_links != len(
    final_toc
):

    print()

    print(
        "❌ 存在没有有效跳转目标的书签。"
    )

    sys.exit(4)


# ============================================================
# 成功
# ============================================================

sys.exit(0)

PYTHON


# ============================================================
# 运行 Python
# ============================================================

echo ""
echo "⑥ 正在定位正文标题..."
echo "   正在排除目录中的同名标题..."
echo "   正在生成 PDF 书签..."
echo ""


python3 "$PY_SCRIPT" \
    "$COPY2" \
    "$PDF1" \
    "$PDF2" \
    "$OUTPUT" \
    "$MAX_OUTLINE_LEVEL"


PY_STATUS=$?


# ============================================================
# 检查 Python
# ============================================================

if [ "$PY_STATUS" -ne 0 ]; then

    echo ""
    echo "=========================================="
    echo " ❌ PDF 书签处理失败"
    echo "=========================================="
    echo ""

    echo "Python 退出代码：$PY_STATUS"
    echo ""

    read -n 1
    exit 1

fi


# ============================================================
# 检查最终文件
# ============================================================

if [ ! -s "$OUTPUT" ]; then

    echo ""
    echo "❌ 最终 PDF 没有正常生成"
    echo ""

    read -n 1
    exit 1

fi


# ============================================================
# 完成
# ============================================================

echo ""
echo "=========================================="
echo " ✅ 博士论文 PDF 生成完成"
echo "=========================================="
echo ""

echo "最终文件："
echo "$OUTPUT"
echo ""

echo "生成结果："
echo "  ✓ Word 打印质量"
echo "  ✓ 高清图片"
echo "  ✓ 自定义样式大纲识别"
echo "  ✓ 自动标题编号"
echo "  ✓ 已排除目录同名标题"
echo "  ✓ PDF 层级书签"
echo "  ✓ PDF 内部页面跳转"
echo "  ✓ 自动校正前置页偏移"
echo "  ✓ 原始 Word 未修改"
echo ""

echo "正在打开最终 PDF..."
echo ""


open "$OUTPUT"


echo ""
echo "按任意键退出..."
read -n 1

exit 0