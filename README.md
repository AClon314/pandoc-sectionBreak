# pandoc-sectionBreak

Rewrite [alexstoick/pandoc-docx-pagebreak](https://github.com/alexstoick/pandoc-docx-pagebreak/blob/master/pandoc-docx-pagebreak.hs) into python.

See [Differences between "???breaks"](https://support.microsoft.com/en-us/office/insert-a-section-break-eef20fd8-e38c-4ba6-a027-e503bdf8375c)

## Install 安装
```sh
pip install git+https://github.com/AClon314/pandoc-sectionBreak.git
```

## Usage 用法
If you want new page have **different** page-number like:

`Ⅰ,Ⅱ,Ⅲ` for Abstract, `1,2,3` for Body, then:

```markdown
---
template: reference-doc.docx
---
摘要

<!-- 分页符: https://github.com/pandoc-ext/pagebreak
会跟随同一节的页边距、页码等格式 -->
\newpage

Abstract end

<!-- 分节符: 独立的页边距、页码等格式 -->
<br section> 

Body start
<br continue> <!-- 连续分节符 -->
<br odd> <!-- 奇数页 -->
<br even> <!-- 偶数页 -->
<br col> <!-- 分列符 -->
```

Assign `template` for using the section format in `<w:sectPr>...<w:sectPr />` block

If you have multiple `<w:sectPr>` in `reference.docx`, you can assign `section: 5` to use the 5th `<w:sectPr>` format. See [test.md](./tests/test.md) for more.

## Dev 调试

If you encounter error, assign environment `DEBUG=1` variable to show debug message:
```sh
DEBUG=1 pandoc --filter pandoc-sectionBreak ...
```

Apply:

- `.html`: css `@media print`
- `.docx`: tag `<w:sectPr>`

See [raw openxml implementation](https://github.com/jgm/pandoc/discussions/10765)
