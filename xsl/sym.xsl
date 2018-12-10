<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:word200x="http://schemas.microsoft.com/office/word/2003/wordml"
  xmlns:v="urn:schemas-microsoft-com:vml" 
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:exsl='http://exslt.org/common'
  xmlns:saxon="http://saxon.sf.net/"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:tr= "http://transpect.io"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:css="http://www.w3.org/1996/css"
  version="2.0" 
  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x docx2hub exsl saxon fn tr mml">

  <xsl:variable name="custom-font-maps" as="document-node(element(symbols))*" select="collection()[symbols]"/>

  <xsl:function name="docx2hub:font-map-name" as="xs:string">
    <xsl:param name="font-map-doc" as="document-node(element(symbols))"/>
    <xsl:sequence select="($font-map-doc/symbols/@docx-name, replace($font-map-doc/*/base-uri(), '^.+/([^/.]+)\.xml', '$1'))[1]"/>
  </xsl:function>

  <xsl:variable name="custom-font-names" as="xs:string*"
    select="for $cfm in $custom-font-maps return docx2hub:font-map-name($cfm)"/>

  <xsl:variable name="docx2hub:symbol-font-names" as="xs:string+" 
    select="('ArialMT+1', 'Math1', 'MT Extra', 'Symbol', 'TimesNewRomanPSMT+1', 'Wingdings', 'Wingdings 2', 
             'Wingdings 3', 'Webdings', 'Euclid Math One', 'Euclid Math Two', 'Euclid Extra', 'Euclid Fraktur', 
             'Lucida Bright Math Italic', 'Lucida Bright Math Extension', 'Lucida Bright Math Symbol', 
             'Monotype Sorts', 'UniversalMath1 BT', 'ZWAdobeF', $custom-font-names)"/>

  <xsl:variable name="docx2hub:symbol-replacement-rfonts" as="element(w:rFonts)">
    <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Cambria Math"/>
  </xsl:variable>
  
  <!-- Set to 'yes' if you want to leave the symbols unaltered for which there exists no Unicode mapping. 
    If set to 'no', they will generate a phrase with the role 'hub:ooxml-symbol' (see below) -->
  <xsl:param name="keep-unmappable-syms" as="xs:string" select="'no'"/>
  
  <xsl:template match="w:sym
                       |
                       w:t[string-length(.)=1 and ../w:rPr/w:rFonts/@w:ascii=$docx2hub:symbol-font-names]
                       |
                       w:t[string-length(.)=1 and ../self::w:r[@role and 
                       (@css:font-family=$docx2hub:symbol-font-names or not(@css:font-family))] and 
                       //css:rule[@layout-type eq 'inline'][@css:font-family=$docx2hub:symbol-font-names]/@name = ../self::w:r/@role]
                       |
                       w:lvlText[../w:rPr/w:rFonts/@w:ascii=$docx2hub:symbol-font-names]/@w:val
                         [if (matches(., '%\d')) then not(../../w:numFmt/@w:val = 'decimal') else true()]
                       |
                       w:t[string-length(.)=1 and ../@css:font-family = $docx2hub:symbol-font-names]
                       |
                       w:lvlText[../@css:font-family = $docx2hub:symbol-font-names]/@w:val
                         [if (matches(., '%\d')) then not(../../w:numFmt/@w:val = 'decimal') else true()]
                       " mode="wml-to-dbk" priority="1.5">
    <!-- priority = 1.5 because of priority = 1 ("default for attributes") in wml2dbk.xsl -->

    <xsl:variable name="font" select="if (self::w:sym) 
                                      then @w:font
                                      else
                                        if (self::attribute(w:val)) (: in w:lvlText :)
                                        then 
                                          if (parent::w:lvlText[../w:rPr/w:rFonts/@w:ascii=$docx2hub:symbol-font-names])
                                          then ../../w:rPr/w:rFonts/@w:ascii
                                          else parent::w:lvlText/../@css:font-family
                                        else 
                                        if(//css:rule[@layout-type eq 'inline'][@css:font-family=$docx2hub:symbol-font-names]/@name = current()/../self::w:r/@role)
                                        then //css:rule[@layout-type eq 'inline'][@name = current()/../self::w:r/@role]/@css:font-family
                                          else (../w:rPr/w:rFonts/@w:ascii, ../@css:font-family)[1]" as="xs:string?"/>
    <xsl:if test="empty($font)">
      <xsl:call-template name="signal-error">
        <xsl:with-param name="error-code" select="'W2D_080'"/>
        <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
        <xsl:with-param name="hash">
          <value key="xpath"><xsl:value-of select="(@srcpath, ancestor::*[@srcpath][1]/@srcpath)[1]"/></value>
          <value key="level">INT</value>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
    <xsl:variable name="number" select="if (self::w:sym) then @w:char else xs:string(.)"/>
    <xsl:variable name="context" select="." as="item()"/>
    <xsl:variable name="font_map" as="document-node(element(symbols))?" select="docx2hub:font-map($font)"/>
    <xsl:variable name="text" as="node()">
      <xsl:choose>
        <xsl:when test="if (self::w:sym) 
                        then $font_map/symbols/symbol[@number = $number]/@char = '&#x000a;' 
                        else $font_map/symbols/symbol[@entity = $number]/@char = '&#x000a;'">
          <br/>
        </xsl:when>
        <xsl:when test="name() eq 'w:val' and matches(., '^%\d')">
          <text>
            <xsl:value-of select="replace($number, '^%\d', '')"/>
          </text>
        </xsl:when>
        <xsl:when test="name() eq 'w:val' and matches(., '^[&#xF000;-&#xF0FF;]%\d$')">
          <!-- https://mantis.le-tex.de/mantis/view.php?id=24633 -->
          <text mapped="true">
            <xsl:value-of select="($font_map/symbols/symbol[@entity = replace($number, '%\d', '')]/@*[name() = (concat('char-', $charmap-policy))], 
                                   $font_map/symbols/symbol[@entity = replace($number, '%\d', '')]/@char)[1]"/>
          </text>
        </xsl:when>
        <xsl:when test="if (self::w:sym) 
                        then $font_map/symbols/symbol[@number = $number] 
                        else $font_map/symbols/symbol[@entity = $number]">
          <text mapped="true">
            <xsl:value-of select="if (self::w:sym) 
                                  then ($font_map/symbols/symbol[@number = $number]/@*[name() = (concat('char-', $charmap-policy))],
                                        $font_map/symbols/symbol[@number = $number]/@char)[1]
                                  else ($font_map/symbols/symbol[@entity = $number]/@*[name() = (concat('char-', $charmap-policy))], 
                                        $font_map/symbols/symbol[@entity = $number]/@char)[1]"/>
          </text>
        </xsl:when>
        <xsl:when test="self::w:sym">
          <xsl:call-template name="create-replacement">
            <xsl:with-param name="font" select="$font"/>
            <xsl:with-param name="number" select="$number"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:message select="'WARNING: Spurious numbering text ', $number, ' in ', $context"/>
          <text>
            <xsl:value-of select="$number"/>
            <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_601', 'WRN', 'wml-to-dbk', 
                                                   concat('Could not map char ', 
                                                          string-join(xs:string(string-to-codepoints($number)), ', '), 
                                                          ' in font ', $font, ' (message c)'
                                                          )
                                                   )"/>
          </text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$text[self::text]">
        <xsl:sequence select="$text/node()"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="not($text//processing-instruction(tr))">
          <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_601', 'WRN', 'wml-to-dbk', 
                                  concat('Could not map char ', (string-to-codepoints($text), $number)[1], ' in font ', $font, ' (message d)'))"/>  
        </xsl:if>
        <xsl:for-each select="$text">
          <xsl:copy>
            <xsl:sequence select="@* except @docx2hub:map-to"/>
            <xsl:sequence select="node()"/>
          </xsl:copy>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="docx2hub:font-map" as="document-node(element(symbols))?">
    <xsl:param name="font-name" as="xs:string?"/>
    <xsl:if test="$font-name">
      <xsl:choose>
        <xsl:when test="$font-name = $custom-font-names">
          <xsl:sequence select="$custom-font-maps[docx2hub:font-map-name(.) = $font-name]"/>  
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="empty($catalog)">
            <xsl:message terminate="yes"
              select="'docx2hub:font-map() in sym.xsl needs a catalog in order to resolve even the most common font maps'"/>
          </xsl:if>
          <xsl:variable name="font-map-name" 
            select="tr:resolve-uri-by-catalog(
                      concat(
                        'http://transpect.io/fontmaps/', 
                        replace($font-name, ' ', '_'),
                        '.xml'
                      ),
                      $catalog
                    )" as="xs:string" />
          <xsl:sequence select="if (doc-available($font-map-name)) then document($font-map-name) else ()"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:function>
  
  <xsl:function name="docx2hub:applied-font-for-w-t" as="xs:string?">
    <xsl:param name="w-t" as="element(w:t)"/>
    <xsl:variable name="rStyle-val" as="xs:string?" select="$w-t/../w:rPr/w:rStyle/@w:val"/>
    <xsl:variable name="pStyle-val" as="xs:string?" select="$w-t/../w:pPr/w:pStyle/@w:val"/>
    <xsl:variable name="rStyle" as="element(w:style)?" 
      select="docx2hub:based-on-chain(key('style-by-id', $rStyle-val, root($w-t)))/w:style[1]"/>
    <xsl:variable name="pStyle" as="element(w:style)?" 
      select="docx2hub:based-on-chain(key('style-by-id', $pStyle-val, root($w-t)))/w:style[1]"/>
    <xsl:variable name="applied-font" as="xs:string?" 
      select="($w-t/../w:rPr/w:rFonts/@w:ascii, $rStyle/w:rPr/w:rFonts/@w:ascii, $pStyle/w:rPr/w:rFonts/@w:ascii)[1]"/>
    <xsl:sequence select="$applied-font"/>
  </xsl:function>

  <xsl:template name="create-replacement">
    <xsl:param name="font" as="xs:string"/><!-- e.g., Wingdings -->
    <xsl:param name="number" as="xs:string"/><!-- hex number, e.g., F064 -->
    <xsl:param name="leave-unmappable-symbols-unchanged" as="xs:boolean?" select="$keep-unmappable-syms = 'yes'" tunnel="yes"/>
    <xsl:choose>
      <xsl:when test="$leave-unmappable-symbols-unchanged">
        <xsl:sequence select="."/>
      </xsl:when>
      <xsl:otherwise>
        <phrase xmlns="http://docbook.org/ns/docbook" role="hub:ooxml-symbol" css:font-family="{$font}" annotations="{$number}"
          srcpath="{(@srcpath, ancestor::*[@srcpath][1]/@srcpath)[1]}">
          <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_601', 'WRN', 'wml-to-dbk', 
                                  concat('Could not map char ', $number, ' in font ', $font, ' (message e)'))"/>
        </phrase>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="create-symbol">
    <xsl:param name="tokens" as="xs:string+"/>
    <xsl:param name="context" as="node()"/>
    <xsl:variable name="font" select="replace($tokens[position()=(index-of($tokens,'\f')+1)], '&quot;', '')" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="count(index-of($tokens,'\f'))=1 and index-of($tokens,'\f') &gt; 1">  
        <xsl:variable name="sym" select="if(index-of($tokens,'\f') eq 2) 
                                         then $tokens[1]
                                         else $tokens[2]" as="xs:string"/>
        <xsl:choose>
          <xsl:when test="$font = $docx2hub:symbol-font-names">
            <xsl:variable name="number" select="if (matches($sym, '^[0-9]+$')) then tr:dec-to-hex(xs:integer($sym)) else 'NaN'"/>
            <xsl:choose>
              <xsl:when test="$number = 'NaN'">
                <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_601', 'WRN', 'wml-to-dbk', 
                        concat('Could not map char ', string-to-codepoints($sym), ' in font ', $font, ' (message b)'))"/>
                <xsl:call-template name="create-replacement">
                  <xsl:with-param name="font" select="$font"/>
                  <xsl:with-param name="number" select="$sym"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="key('symbol-by-number', upper-case($number), $symbol-font-map)/@char"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_601', 'WRN', 'wml-to-dbk', 
                    concat('Could not map char ', string-to-codepoints($sym), ' in font ', $font, ' (message a)'))"/>
            <xsl:call-template name="create-replacement">
              <xsl:with-param name="font" select="$font"/>
              <xsl:with-param name="number" select="$sym"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_602', 'WRN', 'wml-to-dbk', string-join($tokens, ' '))"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="tr:EQ-string-to-unicode" as="xs:string">
    <xsl:param name="input" as="xs:string"/>
    <xsl:variable name="letters" as="xs:string*">
      <xsl:for-each select="string-to-codepoints($input)">
        <xsl:choose>
          <xsl:when test=". ge 61472 and . le 61659">
            <!-- between F020 and F0FF -->
            <xsl:sequence select="string(key('symbol-by-number', upper-case(tr:dec-to-hex(.)), $symbol-font-map)/@char)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:sequence select="codepoints-to-string(.)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>
    <xsl:sequence select="string-join($letters, '')"/>
  </xsl:function>

  <xsl:function name="tr:dec-to-hex" as="xs:string">
    <xsl:param name="in" as="xs:integer?"/>
    <xsl:sequence select="if (not($in) or ($in eq 0)) 
                          then '0' 
                          else concat(
                            if ($in gt 15) 
                            then tr:dec-to-hex($in idiv 16) 
                            else '',
                            substring('0123456789abcdef', ($in mod 16) + 1, 1)
                          )"/>
  </xsl:function>

</xsl:stylesheet>
