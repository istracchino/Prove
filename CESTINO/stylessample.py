import xml.etree.ElementTree as Xet

Paragraph = ''

style = '''
    <w:style w:type="paragraph" w:styleId="Titolo1">
    <w:name w:val="heading 1"/>
    <w:aliases w:val="Procedura"/>
    <w:basedOn w:val="Normale"/>
    <w:next w:val="Normale"/>
    <w:qFormat/>
    <w:rsid w:val="00AC0C30"/>
    <w:pPr>
        <w:keepNext/>
        <w:pageBreakBefore/>
        <w:numPr>
            <w:numId w:val="12"/>
        </w:numPr>
        <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:bCs/>
        <w:color w:val="A4001D"/>
        <w:kern w:val="28"/>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
    </w:rPr></w:style>'''

#<w:p w14:paraId="245A110A" w14:textId="59A63E0B" w:rsidR="00556065" w:rsidRDefault="0071595E" w:rsidP="0071595E"><w:pPr><w:pStyle w:val="Titolo"/><w:jc w:val="center"/></w:pPr><w:r><w:t>$TITOLO</w:t></w:r></w:p><w:p w14:paraId="59A71B57" w14:textId="75384D1E" w:rsidR="0071595E" w:rsidRPr="0071595E" w:rsidRDefault="0071595E" w:rsidP="0071595E"><w:pPr><w:pStyle w:val="Sottotitolo"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:jc w:val="center"/></w:pPr><w:r><w:t>$STEP</w:t></w:r></w:p>

parag = '''<?xml version=\'1.0\' encoding=\'utf8\'?>
<ns0:p xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ns1="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:ns2="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:ns3="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:ns4="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:ns5="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:ns6="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:ns7="http://schemas.microsoft.com/office/drawing/2010/main" ns1:paraId="5FF04B38" ns1:textId="47F8C66D" ns0:rsidR="008A0CC4" ns0:rsidRDefault="001B1A5D" ns0:rsidP="001B1A5D">
<ns0:pPr>
<ns0:pStyle ns0:val="Paragrafoelenco" />
<ns0:numPr>
<ns0:ilvl ns0:val="0" />
<ns0:numId ns0:val="1" />
</ns0:numPr><ns0:jc ns0:val="center" />
</ns0:pPr>
<ns0:r>
<ns0:rPr>
<ns0:noProof />
</ns0:rPr>
<ns0:drawing>
<ns2:anchor distT="0" distB="288290" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" ns3:anchorId="6702BF0A" ns3:editId="31A9D021">
<ns2:simplePos x="0" y="0" />
<ns2:positionH relativeFrom="page"><ns2:align>center</ns2:align></ns2:positionH>
<ns2:positionV relativeFrom="line"><ns2:posOffset>290014</ns2:posOffset></ns2:positionV>
<ns2:extent cx="5400000" cy="5400000" />
<ns2:effectExtent l="0" t="0" r="0" b="0" />
<ns2:wrapTopAndBottom /><ns2:docPr id="1" name="Immagine 1" />
<ns2:cNvGraphicFramePr><ns4:graphicFrameLocks noChangeAspect="1" />
</ns2:cNvGraphicFramePr><ns4:graphic><ns4:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<ns5:pic><ns5:nvPicPr><ns5:cNvPr id="0" name="Picture 1" /><ns5:cNvPicPr preferRelativeResize="0"><ns4:picLocks noChangeAspect="1" noChangeArrowheads="1" />
</ns5:cNvPicPr></ns5:nvPicPr><ns5:blipFill><ns4:blip ns6:embed="rId5"><ns4:extLst><ns4:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><ns7:useLocalDpi val="0" />
</ns4:ext></ns4:extLst></ns4:blip><ns4:srcRect /><ns4:stretch><ns4:fillRect /></ns4:stretch></ns5:blipFill><ns5:spPr bwMode="auto">
<ns4:xfrm><ns4:off x="0" y="0" /><ns4:ext cx="5400000" cy="5400000" /></ns4:xfrm><ns4:prstGeom prst="rect">
<ns4:avLst /></ns4:prstGeom><ns4:noFill /><ns4:ln><ns4:noFill /></ns4:ln></ns5:spPr></ns5:pic></ns4:graphicData></ns4:graphic>
<ns3:sizeRelH relativeFrom="margin"><ns3:pctWidth>0</ns3:pctWidth></ns3:sizeRelH><ns3:sizeRelV relativeFrom="margin"><ns3:pctHeight>0</ns3:pctHeight></ns3:sizeRelV></ns2:anchor>
</ns0:drawing></ns0:r><ns0:r ns0:rsidR="00926D5A"><ns0:t>Allora</ns0:t></ns0:r></ns0:p>
'''

<ns0:p xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ns1="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:ns2="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
xmlns:ns3="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:ns4="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:ns5="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:ns6="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:ns7="http://schemas.microsoft.com/office/drawing/2010/main" ns1:paraId="65582C1A" ns1:textId="59DE9D83" ns0:rsidR="00AC6785" ns0:rsidRPr="00AC6785" ns0:rsidRDefault="00D93B22" ns0:rsidP="000033F8"><ns0:pPr><ns0:pStyle ns0:val="Sottotitolo" /><ns0:numPr><ns0:ilvl ns0:val="0" /><ns0:numId ns0:val="2" /></ns0:numPr><ns0:jc ns0:val="center" /></ns0:pPr><ns0:r><ns0:rPr><ns0:noProof /></ns0:rPr><ns0:drawing><ns2:anchor distT="180340" distB="180340" distL="114300" distR="114300" simplePos="0" relativeHeight="251656704" behindDoc="0" locked="1" layoutInCell="1" allowOverlap="0" ns3:anchorId="75FDE8C4" ns3:editId="159216B0"><ns2:simplePos x="0" y="0" /><ns2:positionH relativeFrom="page"><ns2:align>center</ns2:align></ns2:positionH><ns2:positionV relativeFrom="paragraph"><ns2:posOffset>414020</ns2:posOffset></ns2:positionV><ns2:extent cx="5400000" cy="3477600" /><ns2:effectExtent l="0" t="0" r="0" b="8890" /><ns2:wrapTopAndBottom /><ns2:docPr id="1" name="Immagine 1" /><ns2:cNvGraphicFramePr><ns4:graphicFrameLocks noChangeAspect="1" /></ns2:cNvGraphicFramePr><ns4:graphic><ns4:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><ns5:pic><ns5:nvPicPr><ns5:cNvPr id="1" name="Immagine 1" /><ns5:cNvPicPr preferRelativeResize="0" /></ns5:nvPicPr><ns5:blipFill><ns4:blip ns6:embed="rId5" cstate="print"><ns4:extLst><ns4:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><ns7:useLocalDpi val="0" /></ns4:ext></ns4:extLst></ns4:blip><ns4:stretch><ns4:fillRect /></ns4:stretch></ns5:blipFill><ns5:spPr><ns4:xfrm><ns4:off x="0" y="0" /><ns4:ext cx="5400000" cy="3477600" /></ns4:xfrm><ns4:prstGeom prst="rect"><ns4:avLst /></ns4:prstGeom></ns5:spPr></ns5:pic></ns4:graphicData></ns4:graphic><ns3:sizeRelH relativeFrom="margin"><ns3:pctWidth>0</ns3:pctWidth></ns3:sizeRelH><ns3:sizeRelV relativeFrom="margin"><ns3:pctHeight>0</ns3:pctHeight></ns3:sizeRelV></ns2:anchor></ns0:drawing></ns0:r><ns0:r ns0:rsidR="0071595E"><ns0:t>$STEP</ns0:t></ns0:r><ns0:r><ns0:t>IMG</ns0:t></ns0:r><ns0:r ns0:rsidR="00B03A86"><ns0:t>2</ns0:t></ns0:r></ns0:p>

def step(TESTO,IMGID):
    stepimg = '''<?xml version=\'1.0\' encoding=\'utf8\'?>
        <ns0:p xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ns1="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:ns2="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:ns3="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:ns4="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:ns5="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:ns6="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:ns7="http://schemas.microsoft.com/office/drawing/2010/main" ns1:paraId="02BCAD9A" ns1:textId="625CF84C" ns0:rsidR="00D93B22" ns0:rsidRPr="00D93B22" ns0:rsidRDefault="00D93B22" ns0:rsidP="00D93B22">
        <ns0:pPr>
        <ns0:pStyle ns0:val="Sottotitolo" />
        <ns0:numPr><ns0:ilvl ns0:val="0" />
        <ns0:numId ns0:val="2" />
        </ns0:numPr>
        <ns0:jc ns0:val="center" />
        </ns0:pPr><ns0:r>
        <ns0:rPr>
        <ns0:noProof />
        </ns0:rPr>
        <ns0:drawing>
        <ns2:anchor distT="0" distB="360045" distL="114300" distR="114300" simplePos="0" relativeHeight="251657216" behindDoc="0" locked="1" layoutInCell="1" allowOverlap="0" ns3:anchorId="75FDE8C4" ns3:editId="1E47DF6E">
        <ns2:simplePos x="0" y="0" />
        <ns2:positionH relativeFrom="page"><ns2:align>center</ns2:align></ns2:positionH>
        <ns2:positionV relativeFrom="paragraph"><ns2:posOffset>360045</ns2:posOffset></ns2:positionV>
        <ns2:extent cx="5400000" cy="3477600" /><ns2:effectExtent l="0" t="0" r="0" b="8890" />
        <ns2:wrapTopAndBottom /><ns2:docPr id="1" name="Immagine 1" /><ns2:cNvGraphicFramePr>
        <ns4:graphicFrameLocks noChangeAspect="1" /></ns2:cNvGraphicFramePr>
        <ns4:graphic><ns4:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><ns5:pic><ns5:nvPicPr><ns5:cNvPr id="1" name="Immagine 1" />
        <ns5:cNvPicPr preferRelativeResize="0" /></ns5:nvPicPr><ns5:blipFill>
        <ns4:blip ns6:embed="'''+str(IMGID)+'" cstate="print"><ns4:extLst><ns4:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">''''
        <ns7:useLocalDpi val="0" /></ns4:ext></ns4:extLst></ns4:blip><ns4:stretch><ns4:fillRect/></ns4:stretch></ns5:blipFill><ns5:spPr><ns4:xfrm><ns4:off x="0" y="0" />
        <ns4:ext cx="5400000" cy="3477600" /></ns4:xfrm><ns4:prstGeom prst="rect"><ns4:avLst /></ns4:prstGeom></ns5:spPr></ns5:pic></ns4:graphicData></ns4:graphic><ns3:sizeRelH relativeFrom="margin">
        <ns3:pctWidth>0</ns3:pctWidth></ns3:sizeRelH><ns3:sizeRelV relativeFrom="margin"><ns3:pctHeight>0</ns3:pctHeight></ns3:sizeRelV>
        </ns2:anchor></ns0:drawing></ns0:r><ns0:r ns0:rsidR="0071595E"><ns0:t>'''+str(TESTO)+'</ns0:t></ns0:r></ns0:p>'

    return Xet.fromstring(stepimg)