"""Browser adapter for the photo-log Word builder.

The Word XML templates and caption/table builders are preserved from
Build_Word_Photo_Table_Into_Existing_Word_Structure.py. Browser-specific
functions replace local path prompts and process one image at a time so large
photo folders do not have to be loaded into memory together.
"""

from pathlib import Path
from io import BytesIO
import json
import math
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

from PIL import Image, ImageOps

DEFAULTS = {
    "width_cm": 8.0,
    "height_cm": 6.0,
    "allow_portrait": False,
    "dpi": 450,
    "jpeg_quality": 95,
    "first_image_number": 100,
    "first_rel_id_number": 100,
}

BROWSER_STAGE_DIR = Path("/tmp/photo_log_stage")

# Keep the existing landscape Word drawing dimensions exactly as-is.
# Portrait photos use the same displayed height and a 4.5:6 width:height ratio.
LANDSCAPE_WIDTH_EMU = 2876550
LANDSCAPE_HEIGHT_EMU = 2162175
PORTRAIT_WIDTH_CM = 4.5
PORTRAIT_HEIGHT_CM = 6.0
PORTRAIT_WIDTH_EMU = round(LANDSCAPE_HEIGHT_EMU * PORTRAIT_WIDTH_CM / PORTRAIT_HEIGHT_CM)
PORTRAIT_HEIGHT_EMU = LANDSCAPE_HEIGHT_EMU

DOCUMENT_DOT_XML_TEMPLATE = (
    """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
            xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
            xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
            xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
            xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
            xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
            xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
            xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
            xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
            xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
            xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:oel="http://schemas.microsoft.com/office/2019/extlst"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:w16cei="http://schemas.microsoft.com/office/word/2026/wordml/cei"
            xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
            xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
            xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
            xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"
            xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
            xmlns:w16sdtfl="http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock"
            xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14">
            <w:body>
                <w:p w14:paraId="0569DA5F" w14:textId="77777777" w:rsidR="00FF0FA9" w:rsidRDefault="005D027F"
                    w:rsidP="00406788">
                    <w:pPr>
                        <w:pStyle w:val="Heading2" />
                    </w:pPr>
                    <w:bookmarkStart w:id="0" w:name="_Toc144475209" />
                    <w:bookmarkStart w:id="1" w:name="_Toc202963503" />
                    <w:r>
                        <w:t xml:space="preserve">Photo </w:t>
                    </w:r>
                    <w:r w:rsidR="00675BD8">
                        <w:t>Table</w:t>
                    </w:r>
                    <w:r w:rsidR="001076D5">
                        <w:t>: Multiple</w:t>
                    </w:r>
                    <w:r w:rsidR="001249DB">
                        <w:t xml:space="preserve"> </w:t>
                    </w:r>
                    <w:r w:rsidR="001076D5">
                        <w:t>Photos</w:t>
                    </w:r>
                    <w:bookmarkEnd w:id="0" />
                    <w:bookmarkEnd w:id="1" />
                </w:p>
                <w:tbl>
                    <w:tblPr>
                        <w:tblStyle w:val="TableGrid" />
                        <w:tblW w:w="5000" w:type="pct" />
                        <w:tblLayout w:type="fixed" />
                        <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1" />
                    </w:tblPr>
                    <w:tblGrid>
                        <w:gridCol w:w="5038" />
                        <w:gridCol w:w="5038" />
                    </w:tblGrid>
                        {PHOTO_TABLE_ROWS}
                </w:tbl>
                <w:p w14:paraId="7E553C89" w14:textId="30D275F1" w:rsidR="006B020E" w:rsidRPr="00E53FB9"
                    w:rsidRDefault="006B020E" w:rsidP="00E53FB9">
                    <w:pPr>
                        <w:pStyle w:val="Normal" />
                    </w:pPr>
                </w:p>
                <w:sectPr w:rsidR="006B020E" w:rsidRPr="00E53FB9" w:rsidSect="00605FBC">
                    <w:headerReference w:type="default" r:id="rId13" />
                    <w:footerReference w:type="default" r:id="rId14" />
                    <w:pgSz w:w="12240" w:h="15840" w:code="1" />
                    <w:pgMar w:top="1440" w:right="1077" w:bottom="799" w:left="1077" w:header="709" w:footer="709"
                        w:gutter="0" />
                    <w:cols w:space="708" />
                    <w:docGrid w:linePitch="360" />
                </w:sectPr>
            </w:body>
        </w:document>
    """
)

DOCUMENT_DOT_XML_DOT_RELS_TEMPLATE = (
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
            Target="endnotes.xml" />
        <Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            Target="header1.xml" />
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
            Target="numbering.xml" />
        <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
            Target="footnotes.xml" />
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
            Target="../customXml/item2.xml" />
        <Relationship Id="rId16" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
            Target="theme/theme1.xml" />
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
            Target="../customXml/item1.xml" />
        <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
            Target="webSettings.xml" />
        <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
            Target="settings.xml" />
        <Relationship Id="rId15" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
            Target="fontTable.xml" />
        <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
            Target="styles.xml" />
        <Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
            Target="footer1.xml" />
        {RELS_PHOTO_RELATIONSHIPS}
    </Relationships>
    """
)

PHOTO_RELS_TEMPLATE = (
        """
        <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        Target="media/{IMAGE_NAME}" />
        """
)

PHOTO_TABLE_ROW_TEMPLATE = (
    r"""
    <w:tr w:rsidR="00C90438" w:rsidRPr="003551CF" w14:paraId="1CD94DAD" w14:textId="77777777" w:rsidTr="005A1BE6">
    <w:trPr> <w:trHeight w:hRule="exact" w:val="3459" /> </w:trPr>
    <w:tc>
        <w:tcPr><w:tcW w:w="2500" w:type="pct" /><w:vAlign w:val="bottom" /></w:tcPr>
        <w:p w14:paraId="48EB1F24" w14:textId="3A8406D3" w:rsidR="00C90438" w:rsidRPr="00974226" w:rsidRDefault="002F100E" w:rsidP="00243B36">
            <w:pPr>
                <w:pStyle w:val="markup" />
                <w:jc w:val="center" />
                <w:rPr>
                    <w:rFonts w:eastAsiaTheme="minorHAnsi" />
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr><w:noProof /></w:rPr>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="522C835A"
                        wp14:editId="49075CE8">
                        <wp:extent cx="{PHOTO_WIDTH_EMU_1}" cy="{PHOTO_HEIGHT_EMU_1}" />
                        <wp:effectExtent l="0" t="0" r="0" b="9525" />
                        <wp:docPr id="1854359591" name="Picture 8" />
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" />
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="0" name="Picture 43" />
                                        <pic:cNvPicPr>
                                            <a:picLocks noChangeAspect="1" noChangeArrowheads="1" />
                                        </pic:cNvPicPr>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="{PHOTO_REL_ID_1}">
                                            <a:extLst>
                                                <a:ext uri="{{28A0092B-C50C-407E-A947-70E740481C1C}}">
                                                    <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0" />
                                                </a:ext>
                                            </a:extLst>
                                        </a:blip>
                                        <a:srcRect />
                                        <a:stretch>
                                            <a:fillRect />
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr bwMode="auto">
                                        <a:xfrm>
                                            <a:off x="0" y="0" />
                                            <a:ext cx="{PHOTO_WIDTH_EMU_1}" cy="{PHOTO_HEIGHT_EMU_1}" />
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst />
                                        </a:prstGeom>
                                        <a:noFill />
                                        <a:ln>
                                            <a:noFill />
                                        </a:ln>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
    </w:tc>
    <w:tc>
        <w:tcPr>
            <w:tcW w:w="2500" w:type="pct" />
            <w:vAlign w:val="bottom" />
        </w:tcPr>
        <w:p w14:paraId="7F167D7E" w14:textId="68DD3E93" w:rsidR="00C90438" w:rsidRPr="00974226"
            w:rsidRDefault="002F100E" w:rsidP="00243B36">
            <w:pPr>
                <w:pStyle w:val="markup" />
                <w:jc w:val="center" />
                <w:rPr>
                    <w:rFonts w:eastAsiaTheme="minorHAnsi" />
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:noProof />
                </w:rPr>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="3DCFE195"
                        wp14:editId="1608C0E8">
                        <wp:extent cx="{PHOTO_WIDTH_EMU_2}" cy="{PHOTO_HEIGHT_EMU_2}" />
                        <wp:effectExtent l="0" t="0" r="0" b="9525" />
                        <wp:docPr id="1196090353" name="Picture 9" />
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks
                                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                noChangeAspect="1" />
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic
                                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="0" name="Picture 45" />
                                        <pic:cNvPicPr>
                                            <a:picLocks noChangeAspect="1" noChangeArrowheads="1" />
                                        </pic:cNvPicPr>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="{PHOTO_REL_ID_2}"> 
                                            <a:extLst>
                                                <a:ext uri="{{28A0092B-C50C-407E-A947-70E740481C1C}}">
                                                    <a14:useLocalDpi
                                                        xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                                                        val="0" />
                                                </a:ext>
                                            </a:extLst>
                                        </a:blip>
                                        <a:srcRect />
                                        <a:stretch>
                                            <a:fillRect />
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr bwMode="auto">
                                        <a:xfrm>
                                            <a:off x="0" y="0" />
                                            <a:ext cx="{PHOTO_WIDTH_EMU_2}" cy="{PHOTO_HEIGHT_EMU_2}" />
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst />
                                        </a:prstGeom>
                                        <a:noFill />
                                        <a:ln>
                                            <a:noFill />
                                        </a:ln>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
    </w:tc>
    </w:tr>
    <w:tr w:rsidR="00C90438" w:rsidRPr="003551CF" w14:paraId="120EB3AD" w14:textId="77777777" w:rsidTr="005A1BE6">
    <w:trPr>
        <w:trHeight w:val="624" />
    </w:trPr>
    <w:tc>
        <w:tcPr>
            <w:tcW w:w="2500" w:type="pct" />
        </w:tcPr>
        <w:p w14:paraId="5785A5E9" w14:textId="4989076E" w:rsidR="00C90438" w:rsidRPr="00EA0E88"
            w:rsidRDefault="00C90438" w:rsidP="00C13E23">
            <w:pPr>
                <w:spacing w:before="60" />
                <w:jc w:val="center" />
            </w:pPr>
            <w:r>
                <w:t xml:space="preserve">Photo </w:t>
            </w:r>
            <w:fldSimple w:instr=" SEQ Photo \* ARABIC ">
                <w:r w:rsidR="005A1BE6">
                    <w:rPr>
                        <w:noProof />
                    </w:rPr>
                    <w:t>1</w:t>
                </w:r>
            </w:fldSimple>
            <w:r>
                <w:t xml:space="preserve"> </w:t>
            </w:r>
            <w:r w:rsidR="001F29E8">
                <w:t>S</w:t>
            </w:r>
            <w:r w:rsidR="001F29E8" w:rsidRPr="00A25485">
                <w:t>torm water piping system</w:t>
            </w:r>
        </w:p>
    </w:tc>
    <w:tc>
        <w:tcPr>
            <w:tcW w:w="2500" w:type="pct" />
        </w:tcPr>
        <w:p w14:paraId="3D3AC119" w14:textId="3C2838AA" w:rsidR="00C90438" w:rsidRPr="00C13E23"
            w:rsidRDefault="00C90438" w:rsidP="00C13E23">
            <w:pPr>
                <w:spacing w:before="60" />
                <w:jc w:val="center" />
            </w:pPr>
            <w:r w:rsidRPr="00C13E23">
                <w:t xml:space="preserve">Photo </w:t>
            </w:r>
            <w:fldSimple w:instr=" SEQ Photo \* ARABIC ">
                <w:r w:rsidR="005A1BE6">
                    <w:rPr>
                        <w:noProof />
                    </w:rPr>
                    <w:t>2</w:t>
                </w:r>
            </w:fldSimple>
            <w:r w:rsidRPr="00C13E23">
                <w:t xml:space="preserve"> </w:t>
            </w:r>
            <w:r w:rsidR="001F29E8">
                <w:t>DHW</w:t>
            </w:r>
            <w:r w:rsidR="001F29E8" w:rsidRPr="00A25485">
                <w:t xml:space="preserve"> piping system</w:t>
            </w:r>
        </w:p>
    </w:tc>
    </w:tr>
    """
)

_CAPTION_TEXT_PATTERN = re.compile(
    r'(<w:r w:rsidR="001F29E8">\s*<w:t(?:\s+xml:space="preserve")?>)'
    r'(.*?)'
    r'(</w:t>\s*</w:r>\s*'
    r'<w:r w:rsidR="001F29E8" w:rsidRPr="00A25485">\s*'
    r'<w:t(?:\s+xml:space="preserve")?>)'
    r'(.*?)'
    r'(</w:t>\s*</w:r>)',
    re.S,
)

def sentence_case(text: str) -> str:
    """
    Convert the CSV description to simple sentence case without otherwise
    rewriting/correcting the wording.

    Example:
        "Storm Water Piping System" -> "Storm water piping system"
        "DHW Piping System"         -> "Dhw piping system"
    """
    text = text.strip()
    if not text:
        return ""
    return text[:1].upper() + text[1:].lower()

def _replace_caption_text(row_xml: str, description_1: str, description_2: str) -> str:
    """
    Replace ONLY the two hard-coded description text portions in the caption
    row of PHOTO_TABLE_ROW_TEMPLATE.

    All surrounding caption XML -- including "Photo ", SEQ fields, paragraph
    properties, run properties, etc. -- is left unchanged.

    The template already splits each sample description across two <w:r>
    elements. To preserve that structure, the generated description is also
    split after its first character.
    """
    matches = list(_CAPTION_TEXT_PATTERN.finditer(row_xml))

    if len(matches) != 2:
        raise ValueError(
            "PHOTO_TABLE_ROW_TEMPLATE changed unexpectedly: "
            f"expected 2 caption description blocks, found {len(matches)}."
        )

    descriptions = [sentence_case(description_1), sentence_case(description_2)]

    # Replace from right to left so string offsets stay valid.
    for match, description in reversed(list(zip(matches, descriptions))):
        if description:
            first_char = escape(description[0])
            remaining_text = escape(description[1:])
        else:
            first_char = ""
            remaining_text = ""

        replacement = (
            match.group(1)
            + first_char
            + match.group(3)
            + remaining_text
            + match.group(5)
        )

        row_xml = (
            row_xml[:match.start()]
            + replacement
            + row_xml[match.end():]
        )

    return row_xml

def _blank_second_cell(row_block: str) -> str:
    """
    Keep the second table cell and its formatting, but remove its paragraph
    contents. This leaves the second column visually empty for an odd final
    photo while retaining the table/cell/paragraph properties from the
    original template.
    """
    cells = list(re.finditer(r"<w:tc>.*?</w:tc>", row_block, flags=re.S))

    if len(cells) < 2:
        raise ValueError(
            "PHOTO_TABLE_ROW_TEMPLATE changed unexpectedly: "
            "expected 2 cells in each row."
        )

    second_cell_match = cells[1]
    second_cell = second_cell_match.group(0)

    paragraph_match = re.search(
        r"<w:p\b[^>]*>.*?</w:p>",
        second_cell,
        flags=re.S,
    )

    if paragraph_match is None:
        raise ValueError(
            "PHOTO_TABLE_ROW_TEMPLATE changed unexpectedly: "
            "second cell does not contain a paragraph."
        )

    paragraph = paragraph_match.group(0)

    opening_tag_match = re.match(r"<w:p\b[^>]*>", paragraph)
    if opening_tag_match is None:
        raise ValueError("Could not locate the paragraph opening tag.")

    ppr_match = re.search(r"<w:pPr>.*?</w:pPr>", paragraph, flags=re.S)

    empty_paragraph = opening_tag_match.group(0)

    # Preserve the existing paragraph properties, if present.
    if ppr_match is not None:
        empty_paragraph += ppr_match.group(0)

    empty_paragraph += "</w:p>"

    blank_cell = (
        second_cell[:paragraph_match.start()]
        + empty_paragraph
        + second_cell[paragraph_match.end():]
    )

    return (
        row_block[:second_cell_match.start()]
        + blank_cell
        + row_block[second_cell_match.end():]
    )

def _blank_second_column(row_xml: str) -> str:
    """
    PHOTO_TABLE_ROW_TEMPLATE contains two <w:tr> elements:
      1. the image row
      2. the caption row

    Blank the second cell in BOTH rows for an odd final photo.
    """
    row_count = 0

    def replace_row(match: re.Match) -> str:
        nonlocal row_count
        row_count += 1
        return _blank_second_cell(match.group(0))

    result = re.sub(
        r"<w:tr\b.*?</w:tr>",
        replace_row,
        row_xml,
        flags=re.S,
    )

    if row_count != 2:
        raise ValueError(
            "PHOTO_TABLE_ROW_TEMPLATE changed unexpectedly: "
            f"expected 2 table rows, found {row_count}."
        )

    return result

def build_photo_relationship(rel_id: str, image_name: str) -> str:
    """
    PHOTO_RELS_TEMPLATE intentionally remains as supplied, with rId100.
    Each generated relationship must nevertheless have a unique rId so the
    corresponding <a:blip r:embed="..."> can identify a specific image.
    """
    relationship_xml = PHOTO_RELS_TEMPLATE.format(IMAGE_NAME=image_name)

    old_id = 'Id="rId100"'
    new_id = f'Id="{rel_id}"'

    if old_id not in relationship_xml:
        raise ValueError(
            'PHOTO_RELS_TEMPLATE changed unexpectedly: Id="rId100" was not found.'
        )

    return relationship_xml.replace(old_id, new_id, 1)

def _word_extent_for_photo(photo_entry: dict) -> tuple[int, int]:
    """Return the Word drawing size for one processed photo.

    Landscape output preserves the original template dimensions. Portrait output
    keeps the exact same displayed height and narrows the image to the requested
    4.5 cm : 6 cm aspect ratio.
    """
    if str(photo_entry.get("orientation", "landscape")).lower() == "portrait":
        return PORTRAIT_WIDTH_EMU, PORTRAIT_HEIGHT_EMU
    return LANDSCAPE_WIDTH_EMU, LANDSCAPE_HEIGHT_EMU


def build_photo_table_rows(photo_entries: list[dict]) -> str:
    """
    Build one PHOTO_TABLE_ROW_TEMPLATE block per pair of photos.
    """
    generated_rows = []

    for index in range(0, len(photo_entries), 2):
        photo_1 = photo_entries[index]
        photo_2 = photo_entries[index + 1] if index + 1 < len(photo_entries) else None

        width_1, height_1 = _word_extent_for_photo(photo_1)
        width_2, height_2 = _word_extent_for_photo(photo_2 or {})

        row_xml = PHOTO_TABLE_ROW_TEMPLATE.format(
            PHOTO_REL_ID_1=photo_1["rel_id"],
            PHOTO_REL_ID_2=photo_2["rel_id"] if photo_2 else "",
            PHOTO_WIDTH_EMU_1=width_1,
            PHOTO_HEIGHT_EMU_1=height_1,
            PHOTO_WIDTH_EMU_2=width_2,
            PHOTO_HEIGHT_EMU_2=height_2,
        )

        row_xml = _replace_caption_text(
            row_xml,
            photo_1["description"],
            photo_2["description"] if photo_2 else "",
        )

        if photo_2 is None:
            row_xml = _blank_second_column(row_xml)

        generated_rows.append(row_xml)

    return "\n".join(generated_rows)

def _number(value, label, minimum, maximum, integer=False):
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{label} must be a number.") from None
    if not math.isfinite(number) or number < minimum or number > maximum:
        raise ValueError(f"{label} must be between {minimum:g} and {maximum:g}.")
    if integer:
        if number != int(number):
            raise ValueError(f"{label} must be a whole number.")
        return int(number)
    return number


def _bool(value, default=False):
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return bool(value)
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def normalize_settings(settings=None):
    settings = dict(DEFAULTS if settings is None else {**DEFAULTS, **settings})
    return {
        "width_cm": _number(settings.get("width_cm"), "Width", 0.5, 50.0),
        "height_cm": _number(settings.get("height_cm"), "Height", 0.5, 50.0),
        "allow_portrait": _bool(settings.get("allow_portrait"), False),
        "dpi": _number(settings.get("dpi"), "DPI", 72, 1200, integer=True),
        "jpeg_quality": _number(settings.get("jpeg_quality"), "JPEG quality", 1, 100, integer=True),
        "first_image_number": _number(settings.get("first_image_number"), "First image number", 1, 999999, integer=True),
        "first_rel_id_number": _number(settings.get("first_rel_id_number"), "First relationship ID", 17, 999999, integer=True),
    }


def browser_options():
    settings = normalize_settings(DEFAULTS)
    return {
        "defaults": settings,
        "template_filename": "Word Photo Log Organs - Clean.zip",
        "csv_columns": ["Photo_Description", "File_Name"],
        "supported_output": "JPEG",
        "notes": {
            "crop": "EXIF orientation correction, RGB conversion, and centered ImageOps.fit crop.",
            "captions": "Descriptions are converted to simple sentence case exactly like the source script.",
        },
    }


def browser_clear_stage():
    """Remove processed-image staging from the Pyodide temporary filesystem."""
    if BROWSER_STAGE_DIR.exists():
        shutil.rmtree(BROWSER_STAGE_DIR)
    BROWSER_STAGE_DIR.mkdir(parents=True, exist_ok=True)
    return json.dumps({"cleared": True})


def browser_process_image(input_path, output_path, settings_json):
    """Process one original photo and keep only the resized JPEG in temporary storage.

    With portrait support disabled, behavior is unchanged: every source is
    centre-cropped to the configured landscape dimensions. With portrait support
    enabled, EXIF-corrected portrait sources are instead centre-cropped to
    exactly 4.5 cm wide by 6 cm high.
    """
    settings = normalize_settings(json.loads(settings_json))

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with Image.open(input_path) as image:
        source_format = image.format or "Unknown"
        image = ImageOps.exif_transpose(image)
        source_size = image.size
        source_orientation = "portrait" if image.height > image.width else "landscape"
        preserve_portrait = settings["allow_portrait"] and source_orientation == "portrait"

        target_width_cm = PORTRAIT_WIDTH_CM if preserve_portrait else settings["width_cm"]
        target_height_cm = PORTRAIT_HEIGHT_CM if preserve_portrait else settings["height_cm"]
        width_px = round((target_width_cm / 2.54) * settings["dpi"])
        height_px = round((target_height_cm / 2.54) * settings["dpi"])

        image = image.convert("RGB")
        cropped_image = ImageOps.fit(
            image,
            (width_px, height_px),
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.5),
        )
        cropped_image.save(
            output_path,
            "JPEG",
            quality=settings["jpeg_quality"],
            dpi=(settings["dpi"], settings["dpi"]),
        )

    output_size = output_path.stat().st_size
    return json.dumps({
        "width_px": width_px,
        "height_px": height_px,
        "source_width_px": source_size[0],
        "source_height_px": source_size[1],
        "source_format": source_format,
        "source_orientation": source_orientation,
        "orientation": "portrait" if preserve_portrait else "landscape",
        "display_width_cm": target_width_cm,
        "display_height_cm": target_height_cm,
        "output_bytes": output_size,
    })


def browser_build_xml(request_json):
    request = json.loads(request_json)
    settings = normalize_settings(request.get("settings"))
    entries_in = request.get("entries") or []
    if not entries_in:
        raise ValueError("No successfully processed photos were supplied.")

    photo_entries = []
    image_number = settings["first_image_number"]
    rel_number = settings["first_rel_id_number"]
    for raw in entries_in:
        description = str(raw.get("description") or "").strip()
        source_name = str(raw.get("source_name") or "").strip()
        if not description:
            description = Path(source_name).stem
        photo_entries.append({
            "source_name": source_name,
            "image_name": str(raw.get("image_name") or f"image{image_number}.jpg"),
            "description": description,
            "rel_id": str(raw.get("rel_id") or f"rId{rel_number}"),
            "orientation": "portrait" if str(raw.get("orientation", "landscape")).lower() == "portrait" else "landscape",
        })
        image_number += 1
        rel_number += 1

    photo_table_rows = build_photo_table_rows(photo_entries)
    document_xml = DOCUMENT_DOT_XML_TEMPLATE.format(PHOTO_TABLE_ROWS=photo_table_rows).lstrip()
    photo_relationships = "\n".join(
        build_photo_relationship(entry["rel_id"], entry["image_name"])
        for entry in photo_entries
    )
    document_rels_xml = DOCUMENT_DOT_XML_DOT_RELS_TEMPLATE.format(
        RELS_PHOTO_RELATIONSHIPS=photo_relationships
    ).lstrip()

    ET.fromstring(document_xml)
    ET.fromstring(document_rels_xml)

    return json.dumps({
        "document_xml": document_xml,
        "document_rels_xml": document_rels_xml,
        "photo_count": len(photo_entries),
        "odd_photo_count": bool(len(photo_entries) % 2),
        "entries": photo_entries,
        "settings": settings,
    }, ensure_ascii=False)


def browser_build_docx(template_zip_path, output_docx_path, entries_json, settings_json):
    """Package the finished Word document using Python's zipfile module.

    This intentionally mirrors the original desktop script: the Word package is
    the base, document.xml and document.xml.rels are replaced, processed JPEGs
    are written into word/media, and one final .docx is produced.
    """
    settings = normalize_settings(json.loads(settings_json))
    entries = json.loads(entries_json) or []
    if not entries:
        raise ValueError("No successfully processed photos were supplied.")

    xml_result = json.loads(browser_build_xml(json.dumps({
        "entries": entries,
        "settings": settings,
    })))

    template_zip_path = Path(template_zip_path)
    output_docx_path = Path(output_docx_path)
    if not template_zip_path.exists():
        raise FileNotFoundError("The template ZIP could not be found in the browser runtime.")

    required = {
        "[Content_Types].xml",
        "word/styles.xml",
        "word/header1.xml",
        "word/footer1.xml",
        "word/document.xml",
        "word/_rels/document.xml.rels",
    }

    generated_media = {f"word/media/{entry['image_name']}" for entry in xml_result["entries"]}
    replace_names = {"word/document.xml", "word/_rels/document.xml.rels"} | generated_media

    for entry in xml_result["entries"]:
        staged = BROWSER_STAGE_DIR / entry["image_name"]
        if not staged.exists() or not staged.is_file():
            raise FileNotFoundError(f"Processed image is missing from temporary storage: {entry['image_name']}")

    output_docx_path.parent.mkdir(parents=True, exist_ok=True)
    if output_docx_path.exists():
        output_docx_path.unlink()

    with zipfile.ZipFile(template_zip_path, "r") as source_zip:
        names = set(source_zip.namelist())
        missing = sorted(required - names)
        if missing:
            raise ValueError("The template is missing required file(s): " + ", ".join(missing))
        bad = source_zip.testzip()
        if bad:
            raise ValueError(f"The template ZIP is corrupt at: {bad}")

        with zipfile.ZipFile(
            output_docx_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=6,
        ) as out_zip:
            # Match the original script's fresh-package behavior: copy files,
            # not directory records or filesystem metadata.
            for info in source_zip.infolist():
                name = info.filename
                if info.is_dir() or name in replace_names:
                    continue
                out_zip.writestr(name, source_zip.read(name), compress_type=zipfile.ZIP_DEFLATED)

            out_zip.writestr(
                "word/document.xml",
                xml_result["document_xml"].encode("utf-8"),
                compress_type=zipfile.ZIP_DEFLATED,
            )
            out_zip.writestr(
                "word/_rels/document.xml.rels",
                xml_result["document_rels_xml"].encode("utf-8"),
                compress_type=zipfile.ZIP_DEFLATED,
            )

            for entry in xml_result["entries"]:
                staged = BROWSER_STAGE_DIR / entry["image_name"]
                out_zip.writestr(
                    f"word/media/{entry['image_name']}",
                    staged.read_bytes(),
                    compress_type=zipfile.ZIP_DEFLATED,
                )

    with zipfile.ZipFile(output_docx_path, "r") as check_zip:
        bad = check_zip.testzip()
        if bad:
            raise RuntimeError(f"The generated Word document failed its ZIP integrity check at: {bad}")
        generated_names = set(check_zip.namelist())
        if not required.issubset(generated_names):
            raise RuntimeError("The generated Word document is missing required package files.")
        if not generated_media.issubset(generated_names):
            raise RuntimeError("One or more processed photos were not added to the Word document.")

    return json.dumps({
        "photo_count": xml_result["photo_count"],
        "odd_photo_count": xml_result["odd_photo_count"],
        "output_bytes": output_docx_path.stat().st_size,
    }, ensure_ascii=False)


def browser_dispatch(request_json):
    request = json.loads(request_json)
    action = request.get("action")
    if action == "options":
        result = browser_options()
    elif action == "build_xml":
        return browser_build_xml(request_json)
    elif action == "validate_settings":
        result = {"settings": normalize_settings(request.get("settings"))}
    else:
        raise ValueError(f"Unknown browser action: {action}")
    return json.dumps(result, ensure_ascii=False)
