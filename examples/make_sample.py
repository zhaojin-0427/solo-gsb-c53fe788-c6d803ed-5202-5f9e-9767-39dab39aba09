"""生成 README 示例用的 sample.docx（含修订、过期引用、缺失目标）。"""

import io
import zipfile

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

STYLES = f"""<w:styles {W_NS}>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="10"/></w:numPr></w:pPr>
  </w:style>
</w:styles>"""

NUMBERING = f"""<w:numbering {W_NS}>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="10"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>"""

BODY = f"""<w:document {W_NS}><w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
    <w:bookmarkStart w:id="1" w:name="_TocIntro"/>
    <w:r><w:t>总则</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
    <w:bookmarkStart w:id="2" w:name="_TocScope"/>
    <w:r><w:t>适用范围</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/><w:rPr><w:del w:id="10" w:author="a" w:date="2026-01-01"/></w:rPr></w:pPr>
    <w:del w:id="11" w:author="a" w:date="2026-01-01">
      <w:bookmarkStart w:id="3" w:name="_TocOld"/>
      <w:r><w:delText>已删除的旧节</w:delText></w:r>
      <w:bookmarkEnd w:id="3"/>
    </w:del></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
    <w:bookmarkStart w:id="4" w:name="_TocTerm"/>
    <w:r><w:t>术语与定义</w:t></w:r><w:bookmarkEnd w:id="4"/></w:p>
  <w:p><w:r><w:t>引用示例：</w:t></w:r></w:p>
  <w:p><w:r><w:t>见</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF _TocScope \\w </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1.1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t>（正确）</w:t></w:r></w:p>
  <w:p><w:r><w:t>见</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF _TocTerm \\w </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1.3</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t>（过期：接受视图中应为 1.2）</w:t></w:r></w:p>
  <w:p><w:r><w:t>见</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF _TocOld \\w </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1.2</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t>（目标已删除）</w:t></w:r></w:p>
</w:body></w:document>"""

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("word/document.xml", XML_DECL + BODY)
    z.writestr("word/styles.xml", XML_DECL + STYLES)
    z.writestr("word/numbering.xml", XML_DECL + NUMBERING)

with open("sample.docx", "wb") as f:
    f.write(buf.getvalue())
print("sample.docx written")
