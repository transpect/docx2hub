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
  xmlns:tr="http://transpect.io"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0"
  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn tr mml">

  <xsl:template match="w:endnoteReference" mode="wml-to-dbk">
    <footnote role="endnote">
      <xsl:variable name="id" select="@w:id"/>
      <xsl:apply-templates select="/*/w:endnotes/w:endnote[@w:id = $id]" mode="#current"/>
    </footnote>
  </xsl:template>

  <xsl:template match="w:endnote" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:endnoteRef" mode="wml-to-dbk">
    <!-- setzt die Nummer der Fußnote. Prüfen!! -->
    <xsl:variable name="endnote-num-format" select="/*/w:settings/w:endnotePr/w:numFmt/@w:val" as="xs:string?"/>
    <phrase role="hub:identifier">
          <xsl:variable name="provisional-endnote-number">
        <xsl:number value="(count(preceding::w:endnoteRef) + 1)" 
            format="{
                      if ($endnote-num-format) 
                      then tr:get-numbering-format($endnote-num-format, '') 
                      else '1'
                    }"/>
    </xsl:variable>
      <xsl:variable name="cardinality" select="if (matches($provisional-endnote-number,'^\*†‡§[0-9]+\*†‡§$')) 
                                               then xs:integer(replace($provisional-endnote-number, '^\*†‡§([0-9]+)\*†‡§$', '$1'))
                                               else 0"/>
      <xsl:value-of select="if (matches($provisional-endnote-number,'^\*†‡§[0-9]+\*†‡§$')) 
                            then string-join((for $i 
                                              in (1 to xs:integer(ceiling($cardinality div 4))) 
                                              return substring($provisional-endnote-number,if (($cardinality mod 4) ne 0) then ($cardinality mod 4) else 4,1)),'') 
                            else if (matches($provisional-endnote-number,'^a[a-z]$')) 
                                 then replace($provisional-endnote-number,'^a([a-z])$','$1$1')
                                 else $provisional-endnote-number"/>
    </phrase>
  </xsl:template>

</xsl:stylesheet>