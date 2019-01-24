<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns="http://www.w3.org/2000/svg"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:math="http://www.w3.org/2005/xpath-functions/math"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:svg="http://www.w3.org/2000/svg"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:d2s="http://transpect.io/drawingml2svg"
  exclude-result-prefixes="a d2s docx2hub mc math v w w14 wp wpg wps xs" 
  version="3.0">
  
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" standalone="no"
              doctype-public="-//W3C//DTD SVG 1.1/EN"
              doctype-system="http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd"/>

  <xsl:variable name="d2s:presetShapeDefinitions" as="document-node(element(presetShapeDefinitions))"
                select="doc('presetShapeDefinitions.xml')"/>
    
  <xsl:variable name="d2s:constants" as="document-node(element(a:constants))">
    <xsl:document>
      <constants xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
        <gd name="t" fmla="val 0"/>
        <gd name="l" fmla="val 0"/>
        <gd name="cd8" fmla="val 2700000"/>
        <gd name="cd4" fmla="val 5400000"/>
        <gd name="cd83" fmla="val 8100000"/>
        <gd name="cd2" fmla="val 10800000"/>
        <gd name="cd85" fmla="val 13500000"/>
        <gd name="cd43" fmla="val 16200000"/>
        <gd name="cd87" fmla="val 18900000"/>
      </constants>
    </xsl:document>
  </xsl:variable>

  <xsl:variable name="d2s:pageWidth-fallback" as="xs:integer"
                select="xs:integer(/w:root/w:document/w:body/w:sectPr/w:pgSz/@w:w * 635)"/>
  <xsl:variable name="d2s:pageHeight-fallback" as="xs:integer"
                select="xs:integer(/w:root/w:document/w:body/w:sectPr/w:pgSz/@w:h * 635)"/>
  <xsl:variable name="d2s:marginTop-fallback" as="xs:integer"
                select="xs:integer(/w:root/w:document/w:body/w:sectPr/w:pgMar/@w:top * 635)"/>
  <xsl:variable name="d2s:marginBottom-fallback" as="xs:integer" 
                select="xs:integer(/w:root/w:document/w:body/w:sectPr/w:pgMar/@w:bottom * 635)"/>
  <xsl:variable name="d2s:marginLeft-fallback" as="xs:integer"
                select="xs:integer(/w:root/w:document/w:body/w:sectPr/w:pgMar/@w:left * 635)"/>
  <xsl:variable name="d2s:marginRight-fallback" as="xs:integer"
                select="xs:integer(/w:root/w:document/w:body/w:sectPr/w:pgMar/@w:right * 635)"/>
  <xsl:variable name="d2s:SatzspWidth-fallback" as="xs:integer"
                select="$d2s:pageWidth-fallback - $d2s:marginRight-fallback - $d2s:marginLeft-fallback"/>
  <xsl:variable name="d2s:SatzspHeight-fallback" as="xs:integer"
                select="$d2s:pageHeight-fallback - $d2s:marginTop-fallback - $d2s:marginBottom-fallback"/>
  <xsl:variable name="d2s:linePitch-fallback" as="xs:integer?"
                select="xs:integer(/w:root/w:document/w:body/w:sectPr/w:docGrid/@w:linePitch * 635)"/>
  
  <xsl:variable name="d2s:sec-layout-map-fallback" as="map(xs:string, xs:integer?)">
    <xsl:map>
      <xsl:map-entry key="'pageWidth'" select="$d2s:pageWidth-fallback"/>
      <xsl:map-entry key="'pageHeight'" select="$d2s:pageHeight-fallback"/>
      <xsl:map-entry key="'marginTop'" select="$d2s:marginTop-fallback"/>
      <xsl:map-entry key="'marginBottom'" select="$d2s:marginBottom-fallback"/>
      <xsl:map-entry key="'marginLeft'" select="$d2s:marginLeft-fallback"/>
      <xsl:map-entry key="'marginRight'" select="$d2s:marginRight-fallback"/>
      <xsl:map-entry key="'SatzspWidth'" select="$d2s:pageWidth-fallback - $d2s:marginRight-fallback - $d2s:marginLeft-fallback"/>
      <xsl:map-entry key="'SatzspHeight'" select="$d2s:pageHeight-fallback - $d2s:marginTop-fallback - $d2s:marginBottom-fallback"/>
      <xsl:map-entry key="'linePitch'" select="$d2s:linePitch-fallback"/>        
    </xsl:map>
  </xsl:variable>
  
  <xsl:variable name="d2s:deg2rad" select="math:pi() div 10800000" as="xs:double"/>
  <xsl:variable name="d2s:emu2dpt" select="72 div 914400" as="xs:double"/>

  <xsl:key name="d2s:gd-by-name" match="a:gd" use="@name"/>  
  
  <xsl:template match="w:sectPr" mode="docx2hub:add-props">
    <xsl:variable name="pageWidth"    select="xs:integer(w:pgSz/@w:w * 635)" as="xs:integer"/>
    <xsl:variable name="pageHeight"   select="xs:integer(w:pgSz/@w:h * 635)" as="xs:integer"/>
    <xsl:variable name="marginTop"    select="xs:integer(w:pgMar/@w:top * 635)" as="xs:integer"/>
    <xsl:variable name="marginBottom" select="xs:integer(w:pgMar/@w:bottom * 635)" as="xs:integer"/>
    <xsl:variable name="marginLeft"   select="xs:integer(w:pgMar/@w:left * 635)" as="xs:integer"/>
    <xsl:variable name="marginRight"  select="xs:integer(w:pgMar/@w:right * 635)" as="xs:integer"/>
    <xsl:variable name="linePitch"    select="xs:integer(w:docGrid/@w:linePitch * 635)"/>
    <xsl:variable name="d2s:sec-layout-map" as="map(xs:string, xs:integer?)">
      <xsl:map>
        <xsl:map-entry key="'pageWidth'"    select="$pageWidth"/>
        <xsl:map-entry key="'pageHeight'"   select="$pageHeight"/>
        <xsl:map-entry key="'marginTop'"    select="$marginTop"/>
        <xsl:map-entry key="'marginBottom'" select="$marginBottom"/>
        <xsl:map-entry key="'marginLeft'"   select="$marginLeft"/>
        <xsl:map-entry key="'marginRight'"  select="$marginRight"/>
        <xsl:map-entry key="'SatzspWidth'"  select="$pageWidth - $marginRight - $marginLeft"/>
        <xsl:map-entry key="'SatzspHeight'" select="$pageHeight - $marginTop - $marginBottom"/>
        <xsl:map-entry key="'linePitch'"    select="$linePitch"/>        
      </xsl:map>
    </xsl:variable>
    <xsl:copy>
      <xsl:apply-templates select="@*, node()" mode="#current">
        <xsl:with-param name="d2s:sec-layout-map" select="$d2s:sec-layout-map" 
                        as="map(xs:string, xs:integer?)" tunnel="yes"/>  
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="mc:AlternateContent[mc:Choice/w:drawing/wp:anchor/a:graphic/a:graphicData]" mode="docx2hub:add-props">
    <xsl:param name="d2s:sec-layout-map" as="map(xs:string, xs:integer?)?" tunnel="yes"/>
    <xsl:variable name="element-name" select="if(parent::w:r|parent::w:p) then 'phrase' else 'sidebar'" as="xs:string"/>
    <xsl:apply-templates select="mc:Choice//a:graphic" mode="d2s:default">
      <xsl:with-param name="d2s:sec-layout-map" select="($d2s:sec-layout-map, $d2s:sec-layout-map-fallback)[1]" 
                      as="map(xs:string, xs:integer?)" tunnel="yes"/>
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template match="svg:svg" mode="wml-to-dbk" xmlns="http://docbook.org/ns/docbook">
    <inlinemediaobject><imageobject><imagedata>
      <xsl:copy-of select="."/>
    </imagedata></imageobject></inlinemediaobject>
  </xsl:template>
  
  <xsl:template match="*[a:graphicData]" mode="d2s:default">
    <xsl:param name="d2s:sec-layout-map" as="map(xs:string, xs:integer?)" tunnel="yes"/>
    <svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
      <xsl:variable name="result" as="element(*)*">
        <xsl:apply-templates select="a:graphicData" mode="#current">
          <xsl:with-param name="d2s:sec-layout-map" select="$d2s:sec-layout-map" 
                          as="map(xs:string, xs:integer?)" tunnel="yes"/>
        </xsl:apply-templates>
      </xsl:variable>
      <xsl:call-template name="viewBox">
        <xsl:with-param name="result" select="$result" as="element(*)*"/>
        <xsl:with-param name="d2s:sec-layout-map" select="$d2s:sec-layout-map" 
                        as="map(xs:string, xs:integer?)" tunnel="yes"/>
      </xsl:call-template>
      <title>drawingml2svg</title>
      <g fill="none" stroke="black">
        <rect width="{$d2s:sec-layout-map('pageWidth') * $d2s:emu2dpt}" 
              height="{$d2s:sec-layout-map('pageHeight') * $d2s:emu2dpt}" 
              x="0" y="0" fill="#bee0cd"/>
        <xsl:apply-templates select="$result" mode="d2s:cleanup"/>
      </g>
    </svg>
  </xsl:template>
 <!--********************************
TEMPLATES
********************************-->    
<!-- Determine ViewBox and port -->
  <xsl:template name="viewBox">
    <xsl:param name="result" as="element(*)*"/>
    <xsl:param name="d2s:sec-layout-map" as="map(xs:string, xs:integer?)" tunnel="yes"/>
    <xsl:variable name="width" as="xs:double"
                  select="  ($d2s:sec-layout-map('pageWidth') - min($result/@d2s:min-x) 
                          - ($d2s:sec-layout-map('pageWidth') - max($result/@d2s:max-x) )) * $d2s:emu2dpt"/>
    <xsl:variable name="height" as="xs:double"
                  select="  ($d2s:sec-layout-map('pageHeight') - min($result/@d2s:min-y) 
                          - ($d2s:sec-layout-map('pageHeight') - max($result/@d2s:max-y) )) * $d2s:emu2dpt"/>
    <xsl:attribute name="width" separator=" ">
      <xsl:sequence select="$width"/>
    </xsl:attribute>
    <xsl:attribute name="height" separator=" ">
      <xsl:sequence select="$height"/>
    </xsl:attribute>
    <xsl:attribute name="viewBox" separator=" ">
      <xsl:sequence select="min($result/@d2s:min-x) * $d2s:emu2dpt"/>
      <xsl:sequence select="min($result/@d2s:min-y) * $d2s:emu2dpt"/>
      <xsl:sequence select="$width"/>
      <xsl:sequence select="$height"/>
    </xsl:attribute>
    <xsl:attribute name="preserveAspectRatio">
      <xsl:sequence select="'xMidYMid '"/>
    </xsl:attribute>
  </xsl:template>
 
<!--Generate shapes -->
  <xsl:template match="a:graphicData" mode="d2s:default">
    <xsl:param name="d2s:sec-layout-map" as="map(xs:string, xs:integer?)" tunnel="yes"/>
    <xsl:variable name="this" as="document-node(element(a:graphicData))">
      <xsl:document>
        <xsl:copy-of select="."/>
      </xsl:document>
    </xsl:variable>
    <xsl:variable name="p-before" select="count(//w:p[. &lt;&lt; current()])" as="xs:integer"/>
    <xsl:message select="'PPPPPPPPPP ',$p-before"/>
    <!--Transformation-->
    <xsl:variable name="phi" as="xs:integer">
      <!--HIER: als 60.000 Grad - IN BOGENMAss WIRD ERST BEI NUTZUNG UMGERECHNET Drehung um Mittelpunkt (wps:wsp/wps:spPr/a:xfrm/@rot, 0)[1]-->
      <xsl:choose>
        <xsl:when test="wps:wsp/wps:spPr/a:xfrm/@rot">
          <xsl:sequence select="wps:wsp/wps:spPr/a:xfrm/@rot"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="0"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="flip" as="xs:string">
      <xsl:choose>
        <xsl:when test="wps:wsp/wps:spPr/a:xfrm/@flipV = 1">
          <xsl:sequence select="'flipV'"/>
        </xsl:when>
        <xsl:when test="wps:wsp/wps:spPr/a:xfrm/@flipH = 1">
          <xsl:sequence select="'flipH'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="'noFlip'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!--X Koordinaten -->
    <xsl:variable name="relativeFrom-x" select="../../wp:positionH/@relativeFrom" as="xs:string"/>
    <xsl:variable name="align-x" as="xs:string?"
      select="(../../wp:positionH/wp:align[normalize-space()], ../../wp:positionH/wp:posOffset, '')[1]"/>
    <xsl:variable name="c-x" select="wps:wsp/wps:spPr/a:xfrm/a:ext/@cx" as="xs:integer"/>
    <xsl:variable name="position-x" select="d2s:pos-x($relativeFrom-x, $align-x, $c-x, $d2s:sec-layout-map)" as="xs:integer"/>
    <xsl:variable name="center-x" select="$c-x idiv 2 + $position-x" as="xs:integer"/>
    <!--Y Koordinaten-->
    <xsl:variable name="relativeFrom-y" select="../../wp:positionV/@relativeFrom" as="xs:string"/>
    <xsl:variable name="align-y" as="xs:string?"
      select="(../../wp:positionV/wp:align[normalize-space()], ../../wp:positionV/wp:posOffset, '')[1]"/>
    <xsl:variable name="c-y" select=".//@cy" as="xs:integer"/>
    <xsl:message select="'aaaaaaaaaaaa ', $align-y"></xsl:message>
    <xsl:variable name="position-y" select="d2s:pos-y($relativeFrom-y, $align-y, $c-y, $p-before, $d2s:sec-layout-map)"
                  as="xs:integer"/>
    <xsl:variable name="center-y" select="$c-y idiv 2 + $position-y" as="xs:integer"/>
    <!--Extremwerte-->
    <xsl:variable name="max-local-x"
      select="
        if ($phi gt 0)
        then
          d2s:max-x($phi, $center-x, $c-x, $position-x, $center-y, $c-y, $position-y)
        else
          ($c-x + $position-x)"
      as="xs:double"/>
    <xsl:variable name="max-local-y"
      select="
        if ($phi gt 0)
        then
          d2s:max-y(($phi * $d2s:deg2rad), $center-x, $c-x, $position-x, $center-y, $c-y, $position-y)
        else
          ($c-y + $position-y)"
      as="xs:double"/>
    <xsl:variable name="min-local-x"
      select="
        if ($phi gt 0)
        then
          d2s:min-x(($phi * $d2s:deg2rad), $center-x, $c-x, $position-x, $center-y, $c-y, $position-y)
        else
          ($position-x)"
      as="xs:double"/>
    <xsl:variable name="min-local-y"
      select="
        if ($phi gt 0)
        then
          d2s:min-y(($phi * $d2s:deg2rad), $center-x, $c-x, $position-x, $center-y, $c-y, $position-y)
        else
          ($position-y)"
      as="xs:double"/>

    <!--PRESET GEOMETRY -->
    <xsl:for-each select=".//a:prstGeom">
      <xsl:variable name="preset" as="document-node(element(*))?">
        <xsl:document>
          <xsl:copy-of select="$d2s:presetShapeDefinitions/presetShapeDefinitions/*[name() = current()/@prst]"/>
        </xsl:document>
      </xsl:variable>
      <g class="{@prst}"
        transform="{d2s:transform($position-x, $position-y, $center-x, $center-y, $phi, $flip)}">
        <xsl:attribute name="d2s:min-x" select="$min-local-x"/>
        <xsl:attribute name="d2s:min-y" select="$min-local-y"/>
        <xsl:attribute name="d2s:max-x" select="$max-local-x"/>
        <xsl:attribute name="d2s:max-y" select="$max-local-y"/>
        <xsl:apply-templates select="$d2s:presetShapeDefinitions/presetShapeDefinitions/*[name() = current()/@prst]" mode="d2s:default">
          <xsl:with-param name="lookup-docs" as="document-node(element(*))+" tunnel="yes"
            select="$preset, $d2s:constants"/>
          <xsl:with-param name="xfrm" select="current()/../a:xfrm" as="element(a:xfrm)" tunnel="yes"/>
        </xsl:apply-templates>
      </g>
    </xsl:for-each>

    <!--CUSTOM GEOMETRY -->
    <xsl:for-each select=".//a:custGeom">
      <!--Pfade erstellen-->
      <g transform="{d2s:transform($position-x, $position-y, $center-x, $center-y, $phi, $flip)}">
        <xsl:attribute name="d2s:min-x" select="$min-local-x"/>
        <xsl:attribute name="d2s:min-y" select="$min-local-y"/>
        <xsl:attribute name="d2s:max-x" select="$max-local-x"/>
        <xsl:attribute name="d2s:max-y" select="$max-local-y"/>
        <xsl:apply-templates select="a:pathLst/a:path" mode="d2s:resolve-fmla"/>
      </g>
    </xsl:for-each>
  </xsl:template>
    
    <!--Templates für die einzelnen path commands-->
    
    <xsl:template match="a:pathLst/a:path" mode="d2s:resolve-fmla">
        <path>
            <xsl:attribute name="d">
                <xsl:apply-templates select="*" mode="#current"/>
            </xsl:attribute>
        </path>
    </xsl:template>
    
    <xsl:template match="a:moveTo" mode="d2s:resolve-fmla">
        <xsl:text>M </xsl:text>
        <xsl:call-template name="path-pt"/>
    </xsl:template>
    
    <xsl:template match="a:cubicBezTo" mode="d2s:resolve-fmla">
        <xsl:text>C </xsl:text>
        <xsl:call-template name="path-pt"/>
    </xsl:template>

    <xsl:template match="a:lnTo" mode="d2s:resolve-fmla">
        <xsl:text>L </xsl:text>
        <xsl:call-template name="path-pt"/>
    </xsl:template>
    
    <xsl:template match="a:quadBezTo" mode="d2s:resolve-fmla">
        <xsl:text>Q </xsl:text>
        <xsl:call-template name="path-pt"/>
    </xsl:template>
    
    <xsl:template match="a:arcTo" mode="d2s:resolve-fmla">
        <xsl:variable name="hR" as="xs:integer">
            <xsl:apply-templates select="@hR" mode="d2s:resolve-fmla"/>
        </xsl:variable>
        <xsl:variable name="wR" as="xs:integer">
            <xsl:apply-templates select="@wR" mode="d2s:resolve-fmla"/>
        </xsl:variable>
        <xsl:variable name="a" as="xs:integer">
            <xsl:sequence select="if ($hR ge $wR) then $hR else $wR"/>
        </xsl:variable>
        <xsl:variable name="b" as="xs:integer">
            <xsl:sequence select="if ($hR le $wR) then $hR else $wR"/>
        </xsl:variable>
        <xsl:variable name="swAng-pre"  as="xs:integer"> 
            <xsl:apply-templates select="@swAng" mode="d2s:resolve-fmla"/>
        </xsl:variable>
        <xsl:variable name="swAng" 
            select="if ($swAng-pre = 21600000) then 21599998 
                        else if($swAng-pre lt 0) then (21600000 + $swAng-pre) 
                        else $swAng-pre"/>
        <xsl:variable name="stAng" as="xs:integer">
            <xsl:apply-templates select="@stAng" mode="d2s:resolve-fmla"/>
        </xsl:variable>
        <xsl:variable name="X1" as="xs:double">
            <xsl:sequence select="d2s:ellipsis-x($a, $b, $stAng)"/>
        </xsl:variable>
        <xsl:variable name="X2" as="xs:double">
            <xsl:sequence select="d2s:ellipsis-x($a, $b,    if ($stAng + $swAng gt 21600000) 
                                                            then ($swAng - (21600000 - $stAng)) 
                                                            else $stAng + $swAng)"/>
        </xsl:variable>
        <xsl:variable name="Y1" as="xs:double">
            <xsl:sequence select="d2s:ellipsis-y($a, $b, $stAng)"/>
        </xsl:variable>
        <xsl:variable name="Y2" as="xs:double">
            <xsl:sequence select="d2s:ellipsis-y($a, $b,    if ($stAng + $swAng gt 21600000) 
                                                            then ($swAng - (21600000 - $stAng)) 
                                                            else $stAng + $swAng)"/>
        </xsl:variable>
        <!-- kleines a = relative positionierung -->
        <xsl:text>a </xsl:text>
        <xsl:value-of select="$wR * $d2s:emu2dpt"/>
        <xsl:text> </xsl:text>
        <xsl:value-of select="$hR * $d2s:emu2dpt"/>
        <xsl:text> 0 </xsl:text>
        <xsl:sequence select="if (abs($swAng-pre) gt 10800000) then 1 else 0"/>
        <xsl:text> </xsl:text>
        <xsl:sequence select="if ($swAng-pre gt 0) then 1 else 0"/>
        <xsl:text> </xsl:text>
        <!--relative Positionierung durch x2 - x1 bzw y2 - y1-->
        <xsl:value-of
         select="($X2 - ($X1)) * $d2s:emu2dpt"/>
        <xsl:text> </xsl:text>
        <xsl:value-of
         select="($Y2 - ($Y1)) * $d2s:emu2dpt"/>
        <xsl:text> </xsl:text>
    </xsl:template>
    
    <xsl:template match="a:close" mode="d2s:resolve-fmla">
        <xsl:text>Z </xsl:text>
    </xsl:template>
    
    <!--Pfadpunkte-->
    <xsl:template name="path-pt">
        <xsl:for-each select="a:pt">
            <xsl:variable name="x" as="xs:integer?">
                <xsl:apply-templates select="@x" mode="d2s:resolve-fmla"/>
            </xsl:variable>
            <xsl:variable name="y" as="xs:integer?">
                <xsl:apply-templates select="@y" mode="d2s:resolve-fmla"/>
            </xsl:variable>
            <xsl:value-of select="$x * $d2s:emu2dpt"/>
            <xsl:text> </xsl:text>
            <xsl:value-of select="$y * $d2s:emu2dpt"/>
            <xsl:text> </xsl:text>
        </xsl:for-each>
    </xsl:template>
    
    <xsl:template match="@*[name() = ('x', 'y', 'hR', 'wR', 'stAng', 'swAng')][matches(., '^\p{Ll}')]" mode="d2s:resolve-fmla">
        <xsl:param name="xfrm" as="element(a:xfrm)" tunnel="yes"/>
        <xsl:param name="lookup-docs" as="document-node(element(*))+" tunnel="yes"/>
       <xsl:message select="'gd | resolved ', string(.),'|',d2s:resolve-gd-token(., $lookup-docs, $xfrm)"/>
        <xsl:attribute name="{name()}">
            <xsl:sequence select="d2s:resolve-gd-token(., $lookup-docs, $xfrm)"/>
        </xsl:attribute>
        
    </xsl:template>
    
    <xsl:template match="presetShapeDefinitions/*" mode="d2s:default">
        <xsl:variable name="resolved-pathLst" as="element(a:pathLst)">
            <xsl:apply-templates select="a:pathLst" mode="d2s:resolve-fmla"/>
        </xsl:variable>
        <xsl:sequence select="$resolved-pathLst/*"/>
    </xsl:template>
    
    <!--Alles reinkopieren bis auf d2s Attribute (lokale min/max rauswerfen)-->
    <xsl:template match="node() | @*" mode="d2s:cleanup d2s:resolve-fmla">
        <xsl:copy copy-namespaces="no">
            <xsl:apply-templates select="@* except @d2s:*, node()" mode="#current"/>
        </xsl:copy>
    </xsl:template>
    
    
     <!--Maxima wenn gedreht-->
    <xsl:function name="d2s:max-x" as="xs:double">
        <xsl:param name="phi" as="xs:integer"/>
        <xsl:param name="center-x" as="xs:integer"/>
        <xsl:param name="c-x" as="xs:integer"/>
        <xsl:param name="pos-x" as="xs:integer"/>
        <xsl:param name="center-y" as="xs:integer"/>
        <xsl:param name="c-y" as="xs:integer"/>
        <xsl:param name="pos-y" as="xs:integer"/>
        <xsl:variable name="cos-phi" select="math:cos($phi * $d2s:deg2rad)" as="xs:double"/>
        <xsl:variable name="sin-phi" select="math:sin($phi * $d2s:deg2rad)" as="xs:double"/>
        <xsl:sequence
            select="
                max((
                ($center-x + (($pos-x + $c-x) - $center-x)   * $cos-phi - ($pos-y           - $center-y) * $sin-phi),
                ($center-x + (($pos-x + $c-x) - $center-x)   * $cos-phi - (($pos-y + $c-y)  - $center-y) * $sin-phi),
                ($center-x + ($pos-x          - $center-x)   * $cos-phi - (($pos-y + $c-y)  - $center-y) * $sin-phi),
                ($center-x + ($pos-x          - $center-x)   * $cos-phi - ($pos-y           - $center-y) * $sin-phi)
                ))"
        />
    </xsl:function>

    <xsl:function name="d2s:max-y"  as="xs:double">
        <xsl:param name="phi"/>
        <xsl:param name="mittelpunkt-x" as="xs:integer"/>
        <xsl:param name="c-x" as="xs:integer"/>
        <xsl:param name="position-x" as="xs:integer"/>
        <xsl:param name="mittelpunkt-y" as="xs:integer"/>
        <xsl:param name="c-y" as="xs:integer"/>
        <xsl:param name="position-y" as="xs:integer"/>
        <xsl:variable name="cos-phi" select="math:cos($phi * $d2s:deg2rad)" as="xs:double"/>
        <xsl:variable name="sin-phi" select="math:sin($phi * $d2s:deg2rad)" as="xs:double"/>
        
        <xsl:sequence
            select="
                max((
                ($mittelpunkt-y + (($position-x + $c-x) - $mittelpunkt-x) * $sin-phi + ($position-y - $mittelpunkt-y) * $cos-phi),
                ($mittelpunkt-y + (($position-x + $c-x) - $mittelpunkt-x) * $sin-phi + (($position-y + $c-y) - $mittelpunkt-y) * $cos-phi),
                ($mittelpunkt-y + ($position-x - $mittelpunkt-x) * $sin-phi + (($position-y + $c-y) - $mittelpunkt-y) * $cos-phi),
                ($mittelpunkt-y + ($position-x - $mittelpunkt-x) * $sin-phi + ($position-y - $mittelpunkt-y) * $cos-phi)
                ))"
        />
    </xsl:function>

    <!--Minima wenn gedreht-->
    <xsl:function name="d2s:min-x"  as="xs:double">
        <xsl:param name="phi"/>
        <xsl:param name="mittelpunkt-x" as="xs:integer"/>
        <xsl:param name="c-x" as="xs:integer"/>
        <xsl:param name="position-x" as="xs:integer"/>
        <xsl:param name="mittelpunkt-y" as="xs:integer"/>
        <xsl:param name="c-y" as="xs:integer"/>
        <xsl:param name="position-y" as="xs:integer"/>
        <xsl:variable name="cos-phi" select="math:cos($phi * $d2s:deg2rad)" as="xs:double"/>
        <xsl:variable name="sin-phi" select="math:sin($phi * $d2s:deg2rad)" as="xs:double"/>
        <xsl:sequence
            select="
                min((
                ($mittelpunkt-x - (($position-x + $c-x) - $mittelpunkt-x) * $cos-phi - ($position-y - $mittelpunkt-y) * $sin-phi),
                ($mittelpunkt-x - (($position-x + $c-x) - $mittelpunkt-x) * $cos-phi - (($position-y + $c-y) - $mittelpunkt-y) * $sin-phi),
                ($mittelpunkt-x - ($position-x - $mittelpunkt-x) * $cos-phi - (($position-y + $c-y) - $mittelpunkt-y) * $sin-phi),
                ($mittelpunkt-x - ($position-x - $mittelpunkt-x) * $cos-phi - ($position-y - $mittelpunkt-y) * $sin-phi)
                ))"
        />
    </xsl:function>

    <xsl:function name="d2s:min-y"  as="xs:double">
        <xsl:param name="phi"/>
        <xsl:param name="mittelpunkt-x" as="xs:integer"/>
        <xsl:param name="c-x" as="xs:integer"/>
        <xsl:param name="position-x" as="xs:integer"/>
        <xsl:param name="mittelpunkt-y" as="xs:integer"/>
        <xsl:param name="c-y" as="xs:integer"/>
        <xsl:param name="position-y" as="xs:integer"/>
        <xsl:variable name="cos-phi" select="math:cos($phi * $d2s:deg2rad)" as="xs:double"/>
        <xsl:variable name="sin-phi" select="math:sin($phi * $d2s:deg2rad)" as="xs:double"/>
        <xsl:sequence
            select="
                min((
                ($mittelpunkt-y + (($position-x + $c-x) - $mittelpunkt-x) * $sin-phi + ($position-y - $mittelpunkt-y) * $cos-phi),
                ($mittelpunkt-y + (($position-x + $c-x) - $mittelpunkt-x) * $sin-phi + (($position-y + $c-y) - $mittelpunkt-y) * $cos-phi),
                ($mittelpunkt-y + ($position-x - $mittelpunkt-x) * $sin-phi + (($position-y + $c-y) - $mittelpunkt-y) * $cos-phi),
                ($mittelpunkt-y + ($position-x - $mittelpunkt-x) * $sin-phi + ($position-y - $mittelpunkt-y) * $cos-phi)
                ))"
        />
    </xsl:function>

    <!--ArcTo Punkte auf Ellipse berechnen-->
    <xsl:function name="d2s:ellipsis-x" as="xs:double">
        <xsl:param name="a" as="xs:double"/>
        <xsl:param name="b" as="xs:double"/>
        <xsl:param name="theta" as="xs:double">
            <!-- 1/30000 * pi -->
        </xsl:param>
        <xsl:sequence
            select="
                ($a * $b div math:sqrt($b * $b + $a * $a *
                (let $tmp := math:tan($theta * $d2s:deg2rad)
                return $tmp * $tmp))
                * (if ($theta ge 5400000 and $theta lt 16200000) then -1
                else 1))"
        />
    </xsl:function>

    <xsl:function name="d2s:ellipsis-y" as="xs:double">
        <xsl:param name="a" as="xs:double"/>
        <xsl:param name="b" as="xs:double"/>
        <xsl:param name="theta" as="xs:double"/>
        <xsl:sequence
            select="
                ($a * $b div math:sqrt($a * $a + $b * $b div
                (let $tmp := math:tan($theta * $d2s:deg2rad)
                return $tmp * $tmp))
                * (if ($theta ge 0 and $theta lt 10800000) then 1
                    else -1))"
        />
    </xsl:function>

    <!--Positionierung von Elementen-->
    <xsl:function name="d2s:pos-x" as="xs:integer">
      <xsl:param name="relativeFrom" as="xs:string"/>
      <xsl:param name="align" as="xs:anyAtomicType"/>
      <xsl:param name="cx" as="xs:integer"/>
      <xsl:param name="d2s:sec-layout-map" as="map(xs:string, xs:integer)"/>
        <xsl:choose>
            <xsl:when test="$relativeFrom = 'page'">
                <xsl:choose>
                    <xsl:when test="$align = 'center'">
                      <xsl:sequence select="($d2s:sec-layout-map('pageWidth') idiv 2) - ($cx idiv 2)"/>
                    </xsl:when>
                    <xsl:when test="$align = 'left'">
                        <xsl:sequence select="0"/>
                    </xsl:when>
                    <xsl:when test="$align = 'right'">
                      <xsl:sequence select="$d2s:sec-layout-map('pageWidth') - $cx"/>
                    </xsl:when>
                    <xsl:when test="$align = 'inside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '^\d+$')">
                        <xsl:sequence select="xs:integer($align)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'X: align angabe relativeFrom page nicht zulässig'"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'margin'">
              <xsl:choose>
                    <xsl:when test="$align = 'center'">
                      <xsl:sequence select="($d2s:sec-layout-map('SatzspWidth') idiv 2) + $d2s:sec-layout-map('marginLeft') - ($cx idiv 2)"/>
                    </xsl:when>
                    <xsl:when test="$align = 'right'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageWidth') - $d2s:sec-layout-map('marginRight') - $cx"/>
                    </xsl:when>
                    <xsl:when test="$align = 'left'">
                        <xsl:sequence select="$d2s:sec-layout-map('marginLeft')"/>
                    </xsl:when>
                    <xsl:when test="$align = 'inside'">
                        <xsl:message select="'align inside wird behandelt wie left, nur relevant wenn buchlayout'"/>
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="$d2s:sec-layout-map('marginLeft') + xs:integer($align)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'X: align angabe relativeFrom margin nicht zulässig '"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'leftMargin'">
                <xsl:choose>
                    <xsl:when test="$align = 'right'">
                        <xsl:sequence select="$d2s:sec-layout-map('marginLeft') - $cx"/>
                    </xsl:when>
                    <xsl:when test="$align = 'left'">
                        <xsl:sequence select="0"/>
                    </xsl:when>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="($d2s:sec-layout-map('marginLeft') idiv 2) - ($cx idiv 2)"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="xs:integer($align)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'X: align angabe relativeFrom leftMargin nicht zulässig'"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'rightMargin'">
                <xsl:choose>
                    <xsl:when test="$align = 'right'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageWidth') - $cx"/>
                    </xsl:when>
                    <xsl:when test="$align = 'left'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageWidth') - $d2s:sec-layout-map('marginRight')"/>
                    </xsl:when>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageWidth') - ($d2s:sec-layout-map('marginRight') idiv 2) - ($cx idiv 2)"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        
                        <xsl:sequence select="$d2s:sec-layout-map('pageWidth') - $d2s:sec-layout-map('marginRight') + xs:integer($align)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'X: align angabe relativeFrom rightMargin nciht zulässig'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'insideMargin'">
                <xsl:choose>
                    <xsl:when test="$align = 'left'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'right'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="xs:integer($align)"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'outsideMargin'">
                <xsl:choose>
                    <xsl:when test="$align = 'left'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'right'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'X: align angabe relativeFrom outsideMargin nicht zulässig'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'column'">
                <xsl:message select="'@relativeFrom column wird derzeit so behandelt wie @relativeFrom margin. Unterschied gibt es nur wenn es mehrere Spalten gibt'"/>
                <xsl:choose>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="($d2s:sec-layout-map('SatzspWidth') idiv 2) + $d2s:sec-layout-map('marginLeft') - ($cx idiv 2)"/>
                    </xsl:when>
                    <xsl:when test="$align = 'right'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageWidth') - $d2s:sec-layout-map('marginRight') - $cx"/>
                    </xsl:when>
                    <xsl:when test="$align = 'left'">
                        <xsl:sequence select="$d2s:sec-layout-map('marginLeft')"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="$d2s:sec-layout-map('marginLeft') + xs:integer($align)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'X: align angabe relativeFrom column nicht zulässig '"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'character'">
                <xsl:choose>
                    <xsl:when test="$align = 'left'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'right'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'horizontal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'X: align angabe relativeFrom character nicht zulässig'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
                <xsl:message select="'relativeFrom Angabe für X noch nicht definiert'"/>
                <xsl:sequence select="533400"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>

    <xsl:function name="d2s:pos-y" as="xs:integer">
      <xsl:param name="relativeFrom" as="xs:string"/>
      <xsl:param name="align" as="xs:anyAtomicType"/>
      <xsl:param name="cy" as="xs:integer"/>
      <xsl:param name="pBefore" as="xs:integer"/>
      <xsl:param name="d2s:sec-layout-map" as="map(xs:string, xs:integer)"/>
      <xsl:choose>
            <xsl:when test="$relativeFrom = 'page'">
                <xsl:choose>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="(($d2s:sec-layout-map('pageHeight') idiv 2) - ($cy idiv 2))"/>
                    </xsl:when>
                    <xsl:when test="$align = 'top'">
                        <xsl:sequence select="0"/>
                    </xsl:when>
                    <xsl:when test="$align = 'bottom'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageHeight') - $cy"/>
                    </xsl:when>
                    <xsl:when test="$align = 'inside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '^\d+$')">
                        <xsl:sequence select="xs:integer($align)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'Y: align angabe relativeFrom page nicht zulässig'"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'margin'">
                <xsl:choose>
                    <xsl:when test="$align = 'top'">
                        <xsl:sequence select="$d2s:sec-layout-map('marginTop')"/>
                    </xsl:when>
                    <xsl:when test="$align = 'bottom'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageHeight') - $d2s:sec-layout-map('marginBottom') - $cy"/>
                    </xsl:when>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="($d2s:sec-layout-map('SatzspHeight') idiv 2) - ($cy idiv 2) + $d2s:sec-layout-map('marginTop')"/>
                    </xsl:when>
                    <xsl:when test="$align = 'inside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '^\d+$')">
                        <xsl:sequence select="xs:integer($align) + $d2s:sec-layout-map('marginTop')"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'Y: align angabe relativeFrom page nicht zulässig'"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'insideMargin'">
                <xsl:choose>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'inside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'top'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'bottom'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'Y: align angabe relativeFrom insideMargin nicht zulässig'"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'outsideMargin'">
                <xsl:choose>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'inside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'top'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'bottom'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'Y: align angabe relativeFrom insideMargin nicht zulässig'"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'topMargin'">
                <xsl:choose>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="($d2s:sec-layout-map('marginTop') idiv 2) - ($cy idiv 2)"/>
                    </xsl:when>
                    <xsl:when test="$align = 'inside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'top'">
                        <xsl:sequence select="0"/>
                    </xsl:when>
                    <xsl:when test="$align = 'bottom'">
                        <xsl:sequence select="$d2s:sec-layout-map('marginTop') - $cy"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="xs:integer($align)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'unzulaessige angabe topMargin align Y'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'bottomMargin'">
                <xsl:choose>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageHeight') - ($d2s:sec-layout-map('marginBottom') idiv 2) - ($cy idiv 2)"/>
                    </xsl:when>
                    <xsl:when test="$align = 'inside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'top'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageHeight') - $d2s:sec-layout-map('marginBottom')"/>
                    </xsl:when>
                    <xsl:when test="$align = 'bottom'">
                        <xsl:sequence select="$d2s:sec-layout-map('pageHeight') - $cy"/>
                    </xsl:when>
                    <xsl:when test="matches($align, '\d')">
                        <xsl:sequence select="$d2s:sec-layout-map('pageHeight') - $d2s:sec-layout-map('marginBottom') + xs:integer($align)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'unzulässige Angabe: vertikal align:', $align, 'relativeFrom:', $relativeFrom"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'paragraph'">
                <xsl:choose>
                    <xsl:when test="matches($align, '\d')">
                      <xsl:sequence select="xs:integer($align) + $d2s:sec-layout-map('marginTop') + ($pBefore * $d2s:sec-layout-map('linePitch')) "/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'unzulaessige angabe bottomMargin align Y'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'line'">
                <xsl:choose>
                    <xsl:when test="$align = 'center'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'line'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'outside'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'top'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:when test="$align = 'bottom'">
                        <xsl:sequence select="533400"/>
                        <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:message select="'Y: align angabe relativeFrom insideMargin nicht zulässig'"/>
                        <xsl:sequence select="533400"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
                <xsl:sequence select="533400"/>
                <xsl:message select="'align', $align, 'vertikal relativeFrom', $relativeFrom, 'nicht definiert'"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>

    <!--Transform -->
    <xsl:function name="d2s:transform">
        <xsl:param name="position-x" as="xs:integer"/>
        <xsl:param name="position-y" as="xs:integer"/>
        <xsl:param name="mittelpunkt-x" as="xs:integer"/>
        <xsl:param name="mittelpunkt-y" as="xs:integer"/>
        <xsl:param name="phi"/>
        <xsl:param name="flip"/>
        <!--Rotation-->
        <xsl:if test="$phi gt 0">
            <xsl:text>rotate( </xsl:text>
            <xsl:value-of select="$phi * 0.0006"/>
            <xsl:text> </xsl:text>
            <xsl:value-of select="$mittelpunkt-x * $d2s:emu2dpt"/>
            <xsl:text> </xsl:text>
            <xsl:value-of select="$mittelpunkt-y * $d2s:emu2dpt"/>
            <xsl:text>) </xsl:text>
        </xsl:if>
        <!--Spiegelung-->
        <xsl:if test="$flip = 'flipV' or $flip = 'flipH'">
            <xsl:text>translate(</xsl:text>
            <xsl:value-of select="$mittelpunkt-x * $d2s:emu2dpt"/>
            <xsl:text> </xsl:text>
            <xsl:value-of select="$mittelpunkt-y * $d2s:emu2dpt"/>
            <xsl:text>) </xsl:text>
            <xsl:choose>
                <xsl:when test="$flip = 'flipV'">
                    <xsl:text>scale(1,-1) </xsl:text>
                </xsl:when>
                <xsl:when test="$flip = 'flipH'">
                    <xsl:text>scale(-1,1) </xsl:text>
                </xsl:when>
            </xsl:choose>
            <xsl:text>translate(</xsl:text>
            <xsl:value-of select="$mittelpunkt-x * (-1) * $d2s:emu2dpt"/>
            <xsl:text> </xsl:text>
            <xsl:value-of select="$mittelpunkt-y * (-1) * $d2s:emu2dpt"/>
            <xsl:text>) </xsl:text>
        </xsl:if>
        <!--Einfache Verschiebung an richtige Stelle-->
        <xsl:text>translate(</xsl:text>
        <xsl:value-of select="$position-x * $d2s:emu2dpt"/>
        <xsl:text> </xsl:text>
        <xsl:value-of select="$position-y * $d2s:emu2dpt"/>
        <xsl:text>) </xsl:text>
    </xsl:function>

    <!--Preset Shapes fmla einzelne token auflösen-->
    <xsl:function name="d2s:resolve-gd-token" as="xs:integer">
        <xsl:param name="token" as="xs:anyAtomicType"/> <!-- xs:string or xs:integer -->
        <xsl:param name="lookup-docs" as="document-node(element(*))+"/>
        <xsl:param name="xfrm" as="element(a:xfrm)"/>
        <xsl:choose>
            <xsl:when test="matches(string($token), '^-?[\d.]+$')">
                <xsl:sequence select="xs:integer($token)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'w'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'r'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'wd2' or string($token) = 'hc'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx idiv 2)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'wd4'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx idiv 4)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'wd5'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx idiv 5)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'wd6'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx idiv 6)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'wd8'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx idiv 8)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'wd10'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx idiv 10)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'wd32'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cx idiv 32)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'h' or string($token) = 'b'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cy)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'vc' or string($token) = 'hd2'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cy idiv 2)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'hd3'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cy idiv 3)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'hd4'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cy idiv 4)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'hd5'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cy idiv 5)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'hd6'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cy idiv 6)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'hd8'">
                <xsl:sequence select="xs:integer($xfrm/a:ext/@cy idiv 8)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'ls'">
                <xsl:sequence select="xs:integer(max($xfrm/a:ext/@cy, $xfrm/a:ext/@cx))"/>
            </xsl:when>
            <xsl:when test="string($token) = 'ss'">
                <xsl:sequence select="xs:integer(if($xfrm/a:ext/@cy lt $xfrm/a:ext/@cx) 
                                                    then $xfrm/a:ext/@cy 
                                                    else $xfrm/a:ext/@cx)"/>
            </xsl:when>
            <xsl:when test="string($token) = 'ssd2'">
                <xsl:sequence select="xs:integer(if($xfrm/a:ext/@cy lt $xfrm/a:ext/@cx) 
                                                    then ($xfrm/a:ext/@cy idiv 2)
                                                    else ($xfrm/a:ext/@cx idiv 2))"/>
            </xsl:when>
            <xsl:when test="string($token) = 'ssd4'">
                <xsl:sequence select="xs:integer(if($xfrm/a:ext/@cy lt $xfrm/a:ext/@cx) 
                                                    then ($xfrm/a:ext/@cy idiv 4)
                                                    else ($xfrm/a:ext/@cx idiv 4))"/>
            </xsl:when>
            <xsl:when test="string($token) = 'ssd6'">
                <xsl:sequence select="xs:integer(if($xfrm/a:ext/@cy lt $xfrm/a:ext/@cx) 
                                                    then ($xfrm/a:ext/@cy idiv 6)
                                                    else ($xfrm/a:ext/@cx idiv 6))"/>
            </xsl:when>
            <xsl:when test="string($token) = 'ssd8'">
                <xsl:sequence select="xs:integer(if($xfrm/a:ext/@cy lt $xfrm/a:ext/@cx) 
                                                    then ($xfrm/a:ext/@cy idiv 8)
                                                    else ($xfrm/a:ext/@cx idiv 8))"/>
            </xsl:when>
            <xsl:when test="string($token) = 'ssd16'">
                <xsl:sequence select="xs:integer(if($xfrm/a:ext/@cy lt $xfrm/a:ext/@cx) 
                                                    then ($xfrm/a:ext/@cy idiv 16)
                                                    else ($xfrm/a:ext/@cx idiv 16))"/>
            </xsl:when>
            <xsl:when test="string($token) = 'ssd32'">
                <xsl:sequence select="xs:integer(if($xfrm/a:ext/@cy lt $xfrm/a:ext/@cx) 
                                                    then ($xfrm/a:ext/@cy idiv 32)
                                                    else ($xfrm/a:ext/@cx idiv 32))"/>
            </xsl:when>
            <xsl:when test="matches(string($token), '^\p{Ll}')">
                <xsl:variable name="gd" as="element(a:gd)"
                    select="
                        (for $ld in $lookup-docs
                        return key('d2s:gd-by-name', $token, $ld))[1]"/>
                <xsl:sequence select="d2s:compute-fmla($gd/@fmla, $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:message select="'unknown token ', $token"/>
                <xsl:sequence select="42"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    <!-- Formel berechnen nachdem einzelne Werte aufgelöst wurden-->
    <xsl:function name="d2s:compute-fmla">
        <xsl:param name="fmla" as="xs:string"/>
        <xsl:param name="lookup-docs" as="document-node(element(*))+"/>
        <xsl:param name="xfrm" as="element(a:xfrm)"/>
        <xsl:variable name="tokenized" as="xs:string+" select="tokenize($fmla, '\s+')"/>
        <xsl:variable name="op" as="xs:string" select="$tokenized[1]"/>
        <xsl:variable name="x" as="xs:integer?"
            select="
                if ($tokenized[2])
                then d2s:resolve-gd-token($tokenized[2], $lookup-docs, $xfrm)
                else ()"/>
        <xsl:variable name="y" as="xs:integer?"
            select="
                if ($tokenized[3])
                then d2s:resolve-gd-token($tokenized[3], $lookup-docs, $xfrm)
                else ()"/>
        <xsl:variable name="z" as="xs:integer?"
            select="
                if ($tokenized[4])
                then d2s:resolve-gd-token($tokenized[4], $lookup-docs, $xfrm)
                else ()"/>
        <!--<xsl:message 
            select=
            "'fmla:', $fmla, '| op: ', $op, '| x: ', $x, '| y:', $y, '| z: ', $z, 'resolve'"/>-->
        <xsl:choose>
            <xsl:when test="$op = '*/'">
                <xsl:sequence select="d2s:resolve-gd-token(($x * $y idiv $z), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = '+-'">
                <xsl:sequence
                    select="d2s:resolve-gd-token(($x + $y - $z), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = '+/'">
                <xsl:sequence
                    select="d2s:resolve-gd-token((($x + $y) idiv $z), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = '?:'">
                <xsl:sequence select="d2s:resolve-gd-token((if ($x gt 0) then $y else $z),
                                                                $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'cat2'">
                <xsl:sequence
                    select="d2s:resolve-gd-token(($x * (math:cos(math:atan(($z idiv $y)*$d2s:deg2rad)))),
                        $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'mod'">
                <xsl:sequence
                    select="d2s:resolve-gd-token((math:sqrt(($x * $x) + ($y * $y) + ($z * $z))), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'pin'">
                <xsl:sequence select="d2s:resolve-gd-token((if ($y lt $x) then $x else if ($y gt $z) then $z else $y),
                        $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'sat2'">
                <xsl:sequence select="d2s:resolve-gd-token(($x * math:sin(math:atan(($z idiv $y)*$d2s:deg2rad))),
                        $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'at2'">
                <xsl:sequence select="d2s:resolve-gd-token((math:atan(($y idiv $x)*$d2s:deg2rad)), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'sin'">
                <xsl:sequence select="d2s:resolve-gd-token(($x * math:sin($y * $d2s:deg2rad)), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'tan'">
                <xsl:sequence select="d2s:resolve-gd-token(($x * math:tan($y * $d2s:deg2rad)), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'cos'">
                <xsl:sequence select="d2s:resolve-gd-token(($x * math:cos($y * $d2s:deg2rad)), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'max'">
                <xsl:sequence select="d2s:resolve-gd-token((if ($x gt $y) then $x else $y), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'min'">
                <xsl:sequence select="d2s:resolve-gd-token((if ($x lt $y) then $x else $y), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'abs'">
                <xsl:sequence select="d2s:resolve-gd-token((if ($x lt 0) then (-1 * $x) else $x), $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'val'">
                <xsl:sequence select="d2s:resolve-gd-token($x, $lookup-docs, $xfrm)"/>
            </xsl:when>
            <xsl:when test="$op = 'sqrt'">
                <xsl:sequence select="d2s:resolve-gd-token((math:sqrt($x)), $lookup-docs, $xfrm)"/>
            </xsl:when>
        </xsl:choose>
    </xsl:function>
</xsl:stylesheet>