<?xml version="1.0" encoding="UTF-8"?>
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
  xmlns:tr="http://transpect.io"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0"
  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn tr">

  <xsl:template match="XE[@fldArgs][not(@docx2hub:contains-markup)]" mode="wml-to-dbk" priority="2">
    <xsl:variable name="context" as="element(XE)" select="."/>
    <xsl:apply-templates select="w:r" mode="#current"/>
    <xsl:choose>
      <xsl:when test="matches(@fldArgs,'&quot;(\s*\\[a-z])*\s*(\w+)?\s*$')">       
        <xsl:if test="matches(@fldArgs, '\\[^bfrity&#x201e;&#x201c;]')">
          <xsl:message select="concat('Unexpected index (XE) field function option in ''', @fldArgs, '''')"/>
        </xsl:if>
        <xsl:variable name="see" as="xs:string?" 
                      select="if(matches(@fldArgs, '\\t')) 
                              then replace(@fldArgs, '^.*\\t\s*&quot;(.+?)&quot;(.*$|$)', '$1') 
                              else ()"/>
        <xsl:variable name="temporary-term" as="node()*" select="tr:extract-chars(@fldArgs,'\','\\')"/>
        <xsl:variable name="real-term" as="node()*">
          <xsl:for-each-group select="$temporary-term" group-starting-with="*:text[matches(.,'^\\')]">
           <xsl:choose>
              <xsl:when test="current-group()[1][self::*:text[matches(.,'^\\')]]"/>
              <xsl:otherwise>
                <xsl:for-each select="current-group()">
                  <xsl:choose>
                    <xsl:when test="self::*:text">
                      <xsl:apply-templates select=".//text()" mode="#current"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:apply-templates select="." mode="#current"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each-group>
        </xsl:variable>
        <xsl:variable name="indexterm">
          <indexterm docx2hub:field-function="yes">
            <xsl:copy-of select="@css:*"/>
            <xsl:call-template name="docx2hub:indexterm-attributes">
              <xsl:with-param name="xe" select="."/>
            </xsl:call-template>
            <xsl:for-each select="('primary', 'secondary', 'tertiary')">
              <xsl:call-template name="indexterm-sub">
                <xsl:with-param name="pos" select="position()"/>
                <xsl:with-param name="real-term" select="$real-term"/>
                <xsl:with-param name="elt" select="."/>
              </xsl:call-template>
            </xsl:for-each>
            <xsl:if test="not($see = '') and not(empty($see))">
              <see>
                <xsl:value-of select="tr:rereplace-chars($see)"/>
              </see>
            </xsl:if>
          </indexterm>  
        </xsl:variable>
        <xsl:if test="normalize-space(@fldArgs)">
          <xsl:apply-templates select="$indexterm" mode="index-processing-1"/>            
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_041', 'WRN', 'wml-to-dbk', 
                                               concat('Unexpected index (XE) field function content in ''', @fldArgs, ''''))"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="docx2hub:indexterm-attributes">
    <xsl:param name="xe" as="element(XE)"/>
    <xsl:variable name="type" as="xs:string?" 
                  select="if(matches($xe/@fldArgs, '\\f')) 
                          then replace($xe/@fldArgs, '^.*\\f\s*&quot;?(.+?)&quot;?\s*(\\.*$|$)', '$1')
                          else 
                            if(some $i in tokenize($xe/@fldArgs,':') satisfies matches($i,'Register§§')) 
                            then replace(tokenize($xe/@fldArgs,':')[matches(.,'.*Register§§')],'.*Register§§(.*)$','$1')
                            else ()"/>
    <xsl:variable name="indexterm-attributes" as="attribute()*">
      <xsl:if test="matches($xe/@fldArgs, '\\i')">
        <xsl:attribute name="role" select="'hub:pagenum-italic'"/>
      </xsl:if>
      <xsl:if test="matches($xe/@fldArgs, '\\b')">
        <xsl:attribute name="role" select="'hub:pagenum-bold'"/>
      </xsl:if>
      <xsl:if test="matches($xe/@fldArgs, '\\s')">
        <xsl:attribute name="class" select="'startofrange'"/>
      </xsl:if>
      <xsl:if test="matches($xe/@fldArgs, '\\e')">
        <xsl:attribute name="class" select="'endofrange'"/>
      </xsl:if>
      <xsl:if test="matches(@fldArgs, '\\r')">
        <xsl:variable name="id" as="xs:string" 
          select="tr:rereplace-chars(replace($xe/@fldArgs, '^.*\\r\s*&quot;?\s*(.+?)\s*&quot;?\s*(\\.*$|$)', '$1'))"/>
        <xsl:variable name="bookmark-start" as="element(w:bookmarkStart)*" 
          select="key('docx2hub:bookmarkStart-by-name', ($id, upper-case($id)), root($xe))"/>
        <xsl:choose>
          <xsl:when test="exists($bookmark-start)">
            <xsl:variable name="start-id" as="attribute(xml:id)">
              <xsl:apply-templates select="$bookmark-start/@w:name" mode="bookmark-id"/>
            </xsl:variable>
            <xsl:variable name="end-id" as="attribute(xml:id)">
              <xsl:apply-templates select="$bookmark-start/@w:name" mode="bookmark-id">
                <xsl:with-param name="end" select="true()"/>
              </xsl:apply-templates>
            </xsl:variable>
            <xsl:attribute name="linkends" select="$start-id, $end-id" separator=" "/>
            <!-- Create distinct startofrange/endofrange indexterms at the anchors specified by linkends in the next pass. -->
          </xsl:when>
        </xsl:choose>
      </xsl:if>
      <xsl:if test="not(empty($type))">
        <xsl:attribute name="type" select="tr:rereplace-chars($type)"/>
      </xsl:if>
    </xsl:variable>
    <xsl:for-each-group select="$indexterm-attributes" group-by="name()">
      <xsl:attribute name="{name()}" select="string-join(current-group(), ' ')"/>
    </xsl:for-each-group>
  </xsl:template>
  
  <xsl:template match="XE[@docx2hub:contains-markup]" mode="wml-to-dbk" priority="2">
    <indexterm docx2hub:field-function="yes">
      <xsl:call-template name="docx2hub:indexterm-attributes">
        <xsl:with-param name="xe" select="."/>
      </xsl:call-template>
      <xsl:variable name="primary-etc" as="document-node()">
        <xsl:document>
          <xsl:apply-templates mode="#current"/>
        </xsl:document>
      </xsl:variable>
      <xsl:for-each-group select="$primary-etc/node()" group-starting-with="dbk:sep">
        <xsl:variable name="pos" as="xs:integer" select="position()"/>
        <xsl:element name="{$primary-secondary-etc[$pos]}">
          <xsl:variable name="sortkey-sep" select="current-group()/self::dbk:sortkey[1]" as="element(dbk:sortkey)?"/>
          <xsl:variable name="sortas" as="node()*" select="current-group()[. >> $sortkey-sep]"/>
          <xsl:variable name="term" as="node()*" select="current-group()[not(self::dbk:sep)][not(. >> $sortkey-sep)]"/>
          <xsl:if test="exists(current-group()[1][self::dbk:inlineequation or self::dbk:equation]|$sortas)">
            <xsl:attribute name="sortas" select="string-join(if (exists($sortas)) then $sortas else current-group(), '')"/>
          </xsl:if>
          <xsl:sequence select="$term"/>
        </xsl:element>
      </xsl:for-each-group>
    </indexterm>
  </xsl:template>
  
  <xsl:template match="XE[@docx2hub:contains-markup]/text()" mode="wml-to-dbk">
    <xsl:analyze-string select="." regex="[:;]">
      <xsl:matching-substring>
        <xsl:choose>
          <xsl:when test=". = ':'">
            <sep/>
          </xsl:when>
          <xsl:otherwise>
            <sortkey>;</sortkey>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:matching-substring>
      <xsl:non-matching-substring>
        <xsl:value-of select="."/>
      </xsl:non-matching-substring>
    </xsl:analyze-string>
  </xsl:template>
  
  <xsl:template match="*[@docx2hub:contains-markup]" mode="wml-to-dbk" priority="1.5">
    <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_051', 'WRN', 'wml-to-dbk', 
                            concat('Unexpected markup in ''', name(), ' ', @fldArgs, ' ', ., ''''))"/>
  </xsl:template>
  
  
  
  <!--<xsl:template match="XE[@fldArgs][m:oMath]" 
    mode="wml-to-dbk" priority="2.5">
    <indexterm>
      <primary>
        <xsl:apply-templates mode="#current"/>
      </primary>
    </indexterm>
  </xsl:template>-->

  <xsl:template name="indexterm-sub">
    <xsl:param name="pos" as="xs:integer"/>
    <xsl:param name="real-term" as="node()*"/>
    <xsl:param name="elt" as="xs:string"/>
    <xsl:variable name="real-term-text" select="string-join($real-term,'')" as="xs:string"/>
    <xsl:if test="count(tokenize($real-term-text,':')[not(matches(.,'Register§§'))]) gt $pos - 1">
      <xsl:variable name="processed" as="node()*">
        <xsl:apply-templates select="$real-term" mode="index-processing"/>
      </xsl:variable>
      <xsl:variable name="processed-text" select="string-join($processed, '')" as="xs:string"/>
      <xsl:if test="$processed-text">
        <!--<xsl:element name="{$elt}">
          <xsl:attribute name="sortas" 
            select="normalize-space(replace(tr:rm-last-quot(tokenize($real-term-text,':')[not(matches(.,'Register§§'))][$pos]),'^&quot;',''))"/>
          <xsl:sequence select="$processed"/>
        </xsl:element>-->
        <xsl:variable name="sortkey" as="xs:string?">
          <xsl:analyze-string select="tokenize($real-term-text,':')[not(matches(.,'Register§§'))][$pos]" 
                              regex="^&quot;?\s*(.+?)(;.+?)?\s*&quot;?$">
            <xsl:matching-substring>
              <xsl:sequence select="replace(regex-group(2), '^;', '')"/>
            </xsl:matching-substring>
          </xsl:analyze-string>
        </xsl:variable>
        <xsl:element name="{$elt}">
          <xsl:if test="$sortkey">
            <xsl:attribute name="sortas" select="$sortkey"/>
          </xsl:if>
          <xsl:sequence select="$processed"/>
        </xsl:element>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <xsl:template match="text()[matches(., '^\s*[xX][eE]\s*&quot;.*$')]" mode="index-processing" priority="10">
    <xsl:value-of select="replace(., '^\s*[xX][eE]\s*&quot;(.*?)&quot;?\s*$', '$1')"/>
  </xsl:template>
  
  <xsl:template match="text()[matches(., '^\s*[xX][eE]\s*$')]" mode="index-processing"/>

  <xsl:template match="text()[matches(.,'^\s*&quot;[^\s]+')]" mode="index-processing" priority="+1">
    <xsl:value-of select="replace(., '^\s*&quot;(.*?)&quot;?\s*$', '$1')"/>
  </xsl:template>
  
  <xsl:variable name="primary-secondary-etc" as="xs:string+" select="('primary', 'secondary', 'tertiary', 'quaternary')"/>
  
  <xsl:function name="tr:primary-secondary-tertiary-number" as="xs:integer?">
    <xsl:param name="name" as="xs:string"/>
    <xsl:sequence select="index-of($primary-secondary-etc, $name)"/>
  </xsl:function>
  
  <xsl:template match="dbk:primary | dbk:secondary | dbk:tertiary" mode="index-processing-1" priority="1">
    <xsl:variable name="content" as="node()*">
      <xsl:sequence select="tr:extract-chars(node(),':',':')"/>
    </xsl:variable>
    <xsl:variable name="pst" select="tr:primary-secondary-tertiary-number(local-name())" as="xs:integer"/>
    <xsl:variable name="processed" as="node()*">
      <xsl:for-each-group select="$content" group-starting-with="*:text[matches(.,'^:')]">
        <xsl:variable name="pos" select="position()"/>
        <xsl:choose>
          <xsl:when test="$pos = $pst and not(exists($content[matches(.,'Register§§')]))">
            <xsl:value-of select="normalize-space(tr:rereplace-chars(replace(current-group()[1]/descendant-or-self::text(),'^:[\s&#160;]*','')))"/>
            <xsl:apply-templates select="current-group()[position() gt 1]" mode="index-processing-2"/>
          </xsl:when>
          <xsl:when test="$pos = $pst + 1 and exists($content[matches(.,'Register§§')])">
            <xsl:value-of select="normalize-space(tr:rereplace-chars(replace(current-group()[1]/descendant-or-self::text(),'^:[\s&#160;]*','')))"/>
            <xsl:apply-templates select="current-group()[position() gt 1]" mode="index-processing-2"/>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:variable>
    <xsl:variable name="processed-text" as="xs:string" 
                  select="replace(string-join($processed, ''), '^(.+?)(;.*)?$', '$1')"/>
    <xsl:if test="normalize-space($processed-text)">
      <xsl:copy>
        <xsl:apply-templates select="@* except @sortas" mode="#current"/>
        <xsl:if test="not(@sortas = $processed-text)">
          <xsl:sequence select="@sortas"/>
        </xsl:if>
        <xsl:sequence select="$processed-text"/>
      </xsl:copy>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="*:text | text()" mode="index-processing-2">
    <xsl:value-of select="tr:rereplace-chars(.)"/>
  </xsl:template>
  
  <xsl:function name="tr:rereplace-chars">
    <xsl:param name="context"/>
    <xsl:value-of select="replace(replace($context,'_quot_','&quot;'),'#_-semi-_-colon-_#',':')"/>
  </xsl:function>
  
  <xsl:function name="tr:extract-chars">
    <xsl:param name="context" as="item()*"/><!-- attribute(fldArgs), text(), or element -->
    <xsl:param name="string-char" as="xs:string"/>
    <xsl:param name="regex-char" as="xs:string"/>
    <xsl:for-each select="$context">
      <xsl:choose>
        <xsl:when test="matches(.,$regex-char)">
          <xsl:choose>
            <xsl:when test="self::text() | self::attribute(fldArgs)">
              <xsl:if test="not(tokenize(.,$regex-char)[1]='')">
                <text>
                  <xsl:value-of select="tokenize(.,$regex-char)[1]"/>
                </text>
              </xsl:if>
              <xsl:for-each select="tokenize(.,$regex-char)[position() gt 1]">
                <text>
                  <xsl:value-of select="concat($string-char,.)"/>
                </text>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="self::*">
              <xsl:variable name="element-name" select="local-name(.)"/>
              <xsl:variable name="attributes" select="./@*" as="attribute()*"/>
              <xsl:if test="not(tokenize(./text(),$regex-char)[1]='')">
                <xsl:element name="{$element-name}">
                  <xsl:sequence select="$attributes"/>
                  <xsl:value-of select="tokenize(./text(),$regex-char)[1]"/>
                </xsl:element>
              </xsl:if>
              <xsl:for-each select="tokenize(.,$regex-char)[position() gt 1]">
                <text>
                  <xsl:value-of select="$string-char"/>
                </text>
                <xsl:element name="{$element-name}">
                  <xsl:sequence select="$attributes"/>
                  <xsl:value-of select="."/>
                </xsl:element>
              </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="."/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="."/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:function>
  
</xsl:stylesheet>