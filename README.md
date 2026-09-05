
[中文](./README.md) | [English](./README_en.md)

# Word2PDF Bookmark Generator

**macOS 上基于 Microsoft Word 与 PyMuPDF 的高清 PDF 导出、自动书签生成与
PDF 合并工具**


------------------------------------------------------------------------

## 说明

使用 macOS 上的 Microsoft Word 将毕业论文另存为 PDF 时：
- 如果选择“适合电子分发”，那么可以自动生成书签，但是会压缩图片画质
- 如果选择“适合打印”，那么画质不会压缩，但是不会自动生成书签

这个项目的作用是，在终端中调用 Microsoft Word 对两个 DOCX 文件完成“适合打印”式导出 PDF，合并导出的多个 PDF ，然后使用 PyMuPDF 为合并的 PDF 生成书签。

**注意：**
- 我这里只有两个 DOCX 文件需要转 PDF、合并 PDF、为 PDF 生成书签，因此我只考虑这一种情况。
- 尽管代码会将原始文件复制到 `merge/` 文件夹中再运行，但保险起见请提前备份好原始文件。

## 文件结构：

```text
project/
├── merge_paper.command
├── word1-preface.docx
└── word2-content.docx
```

## 使用方式

- 使用之前，需要安装 PyMuPDF
```shell
python3 -m pip install pymupdf
```

- 将 `merge_paper.command` 文件下载放入`project/` 中

- 在终端中赋予权限
```shell
chmod +x merge_paper.command
```

- 在 `merge_paper.command` 开头修改
```text
WORD1_NAME="word1-preface.docx"
WORD2_NAME="word2-content.docx"
```

- 双击 `merge_paper.command` 运行

- 如果 Microsoft Word 弹框请求权限，需要允许

## 最后生成结果

`merge/` 文件夹中的东西都是代码运行的结果，其中 `Dissertation-Final.pdf` 就是最终版带书签的 PDF 文件

```text
project/
├── merge_paper.command
├── word1-preface.docx
├── word2-content.docx
│
└── merge/
    ├── word1-preface.docx
    ├── word2-content.docx
    ├── word1-preface.pdf
    ├── word2-content.pdf
    ├── create_bookmarks.py
    ├── Dissertation-Final.pdf
    └── Dissertation-Final.pdf.bookmark_log.txt
```

