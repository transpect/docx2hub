<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:docx2hub="http://transpect.io/docx2hub"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    version="2.0"
    exclude-result-prefixes = "xs">

  <!-- The predicate prop[…] is needed when there are multiple entries with different @hubversion
    attributes. The key picks, of all prop declarations that are compatible with the requested $hub-version, 
    the prop declaration for the most recent version. Versions numbers are expected to be dot-separated integers.
    -->
  <xsl:key 
    name="docx2hub:prop" 
    match="prop[
      if(@hubversion)
      then (
        compare($hub-version, @hubversion, 'http://saxon.sf.net/collation?alphanumeric=yes') ge 0
        and @hubversion = max(
          ../prop
            [@name = current()/@name]
            [compare($hub-version, @hubversion, 'http://saxon.sf.net/collation?alphanumeric=yes') ge 0]
              /@hubversion, 
          'http://saxon.sf.net/collation?alphanumeric=yes'
        )
      )
      else true()
    ]"
    use="@name" />

  <xsl:variable name="docx2hub:propmap" as="document-node(element(propmap))">
    <xsl:document xmlns="">
      <propmap>
        <prop name="v:shape/@fillcolor" type="docx-color" target-name="css:background-color"/>
        <prop name="v:shape/@id" />
        <prop name="v:shape/@o:spid" />
        <prop name="v:shape/@o:gfxdata" />
        <prop name="v:shape/@o:allowincell" implement="maybe later" />
        <prop name="v:shape/@o:allowoverlap" implement="maybe later" />
        <prop name="v:shape/@stroked" type="linear" target-name="stroked"/><!-- no passthru b/c no atts may be produced in mode add-atts -->
        <prop name="v:shape/@strokeweight" type="linear" target-name="strokeweight"/>
        <prop name="v:shape/@style" type="linear" target-name="style"/>
        <prop name="v:shape/@type" />
        <prop name="w:adjustRightInd" />
      	<prop name="w:autoRedefine"/>
        <prop name="w:autoSpaceDE" />
        <prop name="w:autoSpaceDN" />
        <prop name="w:b" type="docx-boolean-prop" target-name="css:font-weight" default="normal" active="bold"/>
        <prop name="w:bCs" />
        <prop name="w:bdr" type="docx-bdr" />
        <prop name="w:bidi" type="docx-boolean-prop" target-name="css:direction" default="ltr" active="rtl"/>
        <prop name="w:caps" type="docx-boolean-prop" target-name="css:text-transform" default="none" active="uppercase"/>
        <prop name="w:cantSplit" type="docx-boolean-prop" target-name="css:break-inside" active="avoid"/>
        <prop name="w:cnfStyle" implement="maybe later" />
        <prop name="w:color" type="docx-color" target-name="css:color"/>
        <prop name="w:contextualSpacing" implement="maybe later"/><!-- collapsible spacing between same-style paras, boolean prop. §17.3.1.9 -->
        <prop name="w:dstrike">
          <val eq="false" />
          <val eq="true" target-name="css:text-decoration-line" target-value="line-through"/>
          <val eq="true" target-name="css:text-decoration-style" target-value="double"/>
        </prop>
        <prop name="w:em">
          <val eq="none"/>
        </prop>
        <prop name="wp:extent/@cx" target-name="css:width" type="docx-image-size-attr"/>
        <prop name="wp:extent/@cy" target-name="css:height" type="docx-image-size-attr"/>
        <prop name="w:gridSpan" /><!-- will be calculated by tables.xsl -->
        <prop name="w:highlight" type="docx-color" target-name="css:background-color"/>
        <prop name="w:i" type="docx-boolean-prop" target-name="css:font-style" default="normal" active="italic"/>
        <prop name="w:iCs" />
        <prop name="w:ind/@w:left" type="docx-length-attr" target-name="css:margin-left" />
        <prop name="w:ind/@w:right" type="docx-length-attr" target-name="css:margin-right" />
        <prop name="w:ind/@w:firstLine" type="docx-length-attr" target-name="css:text-indent" />
        <prop name="w:ind/@w:hanging" type="docx-length-attr-negated" target-name="css:text-indent" />
        <prop name="w:jc">
          <val match="left" target-name="css:text-align" target-value="left" />
          <val match="start" target-name="css:text-align" target-value="left" />
          <val match="right" target-name="css:text-align" target-value="right" />
          <val match="end" target-name="css:text-align" target-value="right" />
          <val match="both" target-name="css:text-align" target-value="justify" />
          <val match="center" target-name="css:text-align" target-value="center" />
          <val match="center" target-name="css:text-align-last" target-value="center" />
        </prop>
        <prop name="m:jc">
          <val match="left" target-name="css:text-align" target-value="left" />
          <val match="start" target-name="css:text-align" target-value="left" />
          <val match="right" target-name="css:text-align" target-value="right" />
          <val match="end" target-name="css:text-align" target-value="right" />
          <val match="both" target-name="css:text-align" target-value="justify" />
          <val match="center" target-name="css:text-align" target-value="center" />
          <val match="center" target-name="css:text-align-last" target-value="center" />
        </prop>
      	<prop name="w:keepLines" />
      	<prop name="w:keepNext" type="docx-boolean-prop" default="auto" active="avoid" target-name="css:page-break-after"/>
      	<prop name="w:kern" />
        <prop name="w:kinsoku" />
        <prop name="w:lang" type="lang" target-name="xml:lang" />
        <prop name="w:link" />
        <prop name="w:locked" />
        <prop name="w:name" type="linear" target-name="role" hubversion="1.0"/>
        <prop name="w:name" type="linear" target-name="native-name" hubversion="1.1"/>
        <!--<prop name="w:name" type="style-name" target-name="name" hubversion="1.1"/>-->
        <prop name="w:next" />
        <prop name="w:noProof" />
        <prop name="w:noWrap" type="docx-boolean-prop" target-name="css:white-space" default="normal" active="nowrap"/>
        <prop name="w14:numForm">
          <!-- if there are multiple modifiers (tabular-nums, etc.), we need to establish a type docx-numForm for this prop -->
          <val match="oldStyle" target-name="css:font-variant-numeric" target-value="oldstyle-nums" />
        </prop>
        <prop name="w:numPr" type="passthru" />
        <prop name="w:outline">
          <val eq="true" target-name="css:text-shadow" target-value="1px 0px"/>
          <val eq="false" />
        </prop>
        <prop name="w:outlineLvl" type="docx-hierarchy-level"/>
        <prop name="w:overflowPunct"/>
        <prop name="w:pageBreakBefore" type="docx-boolean-prop" target-name="css:page-break-before" default="auto" active="always"/>
        <prop name="w:pBdr/w:bottom" type="docx-border" />
        <prop name="w:pBdr/w:left" type="docx-border" />
        <prop name="w:pBdr/w:right" type="docx-border" />
        <prop name="w:pBdr/w:top" type="docx-border"  />
        <prop name="w:pgSz/@w:w" type="docx-length-attr" target-name="css:width"/>
        <prop name="w:pgSz/@w:h" type="docx-length-attr" target-name="css:height"/>
        <prop name="w:pgSz/@w:code"/>
        <prop name="w:pgSz/@w:orient" target-name="css:orientation" type="linear"/>
        <prop name="w:position" target-name="css:top" type="docx-position-attr-negated" />
        <prop name="w:position" target-name="css:position" target-value="relative"/>
        <prop name="w:pPrChange"/>
        <prop name="w:pStyle" type="docx-parastyle" />
        <prop name="w:qFormat" />
        <prop name="w:rFonts" type="docx-font-family" target-name="css:font-family" />
        <prop name="w:rPrChange"/>
        <prop name="w:rsid" />
        <prop name="w:rsidDel" /><!-- 17.3.2.25 suggests that it identifies deleted runs. But see for example
            DIN_EN_1673_A1_nf_11758779.doc for two rsidDel runs that are part of the final document. -->
        <prop name="w:rsidR" />
        <prop name="w:rsidRDefault" />
        <prop name="w:rsidP" />
        <prop name="w:rsidRPr" />
        <prop name="w:rStyle" type="docx-charstyle" />
        <prop name="w:rtl" type="docx-boolean-prop" target-name="css:direction" default="ltr" active="rtl"/>
        <prop name="w:rubyAlign">
          <val match="center" target-name="css:ruby-align" target-value="center"/>
          <val match="distributeSpace" target-name="css:ruby-align" target-value="space-around"/>
          <val match="distributeLetter" target-name="css:ruby-align" target-value="space-between"/>
          <val match="left" target-name="css:ruby-position" target-value="left"/>
          <val match="right" target-name="css:ruby-position" target-value="right"/>
          <val match="rightVertical" target-name="css:ruby-position" target-value="inter-character"/>
        </prop>
        <!--<prop name="w:sectPr" />--><!-- provisional -->
        <prop name="w:semiHidden" />
        <prop name="w:shadow" type="docx-boolean-prop" target-name="css:text-shadow" default="none" active="1pt 1pt"/>
        <prop name="w14:shadow" implement="maybe later"/>
        <prop name="w:shd" type="docx-shd" />
        <prop name="w:smallCaps" type="docx-boolean-prop" target-name="css:font-variant" default="normal" active="small-caps"/>
      	<prop name="w:snapToGrid" />
      	<prop name="w:spacing/@w:after" type="docx-length-attr" target-name="css:margin-bottom" />
        <prop name="w:spacing/@w:before" type="docx-length-attr" target-name="css:margin-top" />
        <prop name="w:spacing/@w:afterLines" implement="maybe later" />
        <prop name="w:spacing/@w:beforeLines" implement="maybe later" />
        <prop name="w:spacing/@w:line" type="docx-line" target-name="css:line-height" />
        <prop name="w:spacing/@w:val" type="docx-length-attr" target-name="css:letter-spacing" >
          <!-- GI 2016-11-08: Although contemporary browsers implement letter-spacing like this docx
          property, they really shouldn’t. https://twitter.com/gimsieke/status/796107927916605440 -->
        </prop>
        <prop name="w:strike" target-name="css:text-decoration-line">
          <val eq="true" target-value="line-through"/>
          <val eq="false" />
        </prop>
        <!--<prop name="w:stri-->
        <prop name="w:suppressAutoHyphens" type="docx-boolean-prop" target-name="css:hyphens" default="auto" active="manual"/>
        <prop name="w:suppressLineNumbers" implement="maybe later"/>
        <prop name="w:sz" type="docx-font-size" target-name="css:font-size" />
        <prop name="w:szCs" />
        <prop name="w:tab/@w:leader" type="linear" target-name="leader" />
        <prop name="w:tab/@w:pos" type="docx-length-attr" target-name="horizontal-position" />
        <prop name="w:tab/@w:val">
          <val eq="decimal" target-name="align" target-value="decimal"/>
          <val eq="left" target-name="align" target-value="left" />
          <val eq="center" target-name="align" target-value="center" />
          <val eq="right" target-name="align" target-value="right" />
          <val eq="num" />
          <val eq="clear" target-name="clear" target-value="yes" />
        </prop>
        <prop name="w:tabs" type="tablist" />
        <prop name="w:tblBorders" type="passthru" />
        <prop name="w:tblBorders/w:bottom" type="docx-border" />
        <prop name="w:tblBorders/w:left" type="docx-border" />
        <prop name="w:tblBorders/w:right" type="docx-border" />
        <prop name="w:tblBorders/w:top" type="docx-border"  />
        <prop name="w:tblBorders/w:insideH" type="docx-border"  />
        <prop name="w:tblBorders/w:insideV" type="docx-border"  />
        <prop name="w:tblCellMar" type="passthru" />
        <prop name="w:tblCellMar/w:bottom" type="docx-padding" />
        <prop name="w:tblCellMar/w:left" type="docx-padding" />
        <prop name="w:tblCellMar/w:right" type="docx-padding" />
        <prop name="w:tblCellMar/w:top" type="docx-padding" />
        <prop name="w:tblGrid" type="passthru" />
        <prop name="w:tblInd" type="docx-length-attr" target-name="css:margin-left">
          <!-- postprocess it; should be margin-right if the table is rtl (§ 17.4.51) --> 
        </prop> 
        <prop name="w:tblLayout" type="docx-boolean-prop" target-name="css:table-layout" default="auto" active="fixed"/>
        <prop name="w:tblLook" type="passthru"/>
        <prop name="w:tblPrEx" type="passthru"/>
        <prop name="w:tblStyle" type="docx-parastyle"/>
        <prop name="w:tblW" type="passthru" />
        <prop name="w:tcBorders/w:bottom" type="docx-border" />
        <prop name="w:tcBorders/w:left" type="docx-border" />
        <prop name="w:tcBorders/w:right" type="docx-border" />
        <prop name="w:tcBorders/w:top" type="docx-border"  />
        <prop name="w:tcBorders/w:insideV" type="docx-border"  />
        <prop name="w:tcBorders/w:insideH" type="docx-border"  />
        <prop name="w:tcMar" type="passthru" />
        <prop name="w:tcMar/w:bottom" type="docx-padding" />
        <prop name="w:tcMar/w:left" type="docx-padding" />
        <prop name="w:tcMar/w:right" type="docx-padding" />
        <prop name="w:tcMar/w:top" type="docx-padding" />
        <prop name="w:tcPrChange"/>
        <prop name="w:tcW/@w:w" type="docx-length-attr" target-name="css:width"/>
        <prop name="w:textDirection" type="docx-text-direction"/>
        <prop name="w14:textFill" implement="maybe later"/>
        <prop name="w14:textOutline" implement="maybe later"/>
        <prop name="w:trHeight" type="docx-table-row-height"/>
        <!-- trPr/… -->
        <prop name="w:gridBefore" type="linear" target-name="w:fill-cells-before"/>
        <prop name="w:wBefore/@w:w" type="docx-length-attr" target-name="w:fill-width-before"/>
        <prop name="w:gridAfter" type="linear" target-name="w:fill-cells-after"/>
        <prop name="w:wAfter/@w:w" type="docx-length-attr" target-name="w:fill-width-after"/>
        <prop name="w:u">
          <val match="none"/>
          <val match="dash" target-name="css:text-decoration-style" target-value="dashed"/>
          <val match="dash" target-name="css:text-decoration-line" target-value="underline"/>
          <val match="dotted" target-name="css:text-decoration-style" target-value="dotted"/>
          <val match="dotted" target-name="css:text-decoration-line" target-value="underline"/>
          <val match="double" target-name="css:text-decoration-style" target-value="double"/>
          <val match="double" target-name="css:text-decoration-line" target-value="underline"/>
          <val match="thick" target-name="css:text-decoration-style" target-value="solid"/>
          <val match="thick" target-name="css:text-decoration-line" target-value="underline"/>
          <val match="single" target-name="css:text-decoration-style" target-value="solid"/>
          <val match="single" target-name="css:text-decoration-line" target-value="underline"/>
          <val match="wav" target-name="css:text-decoration-style" target-value="wavy"/>
          <val match="wav" target-name="css:text-decoration-line" target-value="underline"/>
        </prop>
        <prop name="w:u/@w:color" type="docx-color" target-name="css:text-decoration-color"/>
          
        <prop name="w:uiPriority" />
        <prop name="w:unhideWhenUsed" />
        <prop name="w:vAlign">
          <val match="auto" target-name="css:vertical-align" target-value="auto"/>
          <val match="bottom" target-name="css:vertical-align" target-value="bottom"/>
          <val match="baseline" target-name="css:vertical-align" target-value="bottom"/>
          <val match="center" target-name="css:vertical-align" target-value="middle"/>
          <val match="top" target-name="css:vertical-align" target-value="top"/>
        </prop>
        <prop name="w:textAlignment">
          <val match="both" target-name="css:vertical-align" target-value="middle"/>
          <val match="bottom" target-name="css:vertical-align" target-value="bottom"/>
          <val match="center" target-name="css:vertical-align" target-value="middle"/>
          <val match="top" target-name="css:vertical-align" target-value="top"/>
        </prop>
        <prop name="w:vanish" type="docx-boolean-prop" target-name="css:display" default="inherit" active="none"/>
      	<prop name="w:vertAlign" type="docx-position" /><!-- superscript etc. -->
        <prop name="w:vMerge" />
      	<prop name="w:w" type="docx-font-stretch" target-name="css:font-stretch"/>
        <prop name="w:webHidden"/>
        <prop name="w:widowControl" type="docx-boolean-prop" target-name="css:orphans" default="1" active="2"/>
        <prop name="w:widowControl" type="docx-boolean-prop" target-name="css:widows" default="1" active="2"/>
      </propmap>
    </xsl:document>
  </xsl:variable>

</xsl:stylesheet>
