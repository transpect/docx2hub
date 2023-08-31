<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:word200x="http://schemas.microsoft.com/office/word/2003/wordml"
  xmlns:v="urn:schemas-microsoft-com:vml" 
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
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

  <xsl:variable name="comment-reference-style-regex" select="'^(Kommentarzeichen)$'"/>

  <xsl:template match="w:commentReference" mode="wml-to-dbk">
    <xsl:if test="//w:commentRangeEnd[@w:id=current()/@w:id][parent::w:tbl]">
      <xsl:apply-templates select="//w:commentRangeEnd[@w:id=current()/@w:id][parent::w:tbl]" mode="#current"/>
    </xsl:if>
    <xsl:apply-templates select="key('comment-by-id', @w:id)" mode="comment">
      <xsl:with-param name="comment-id" as="xs:string" tunnel="yes" select="@w:id"/>
    </xsl:apply-templates>
  </xsl:template>

  <!-- dissolve single w:r with only comment(s) -->
  <xsl:template match="*[*]
                        [self::w:r[matches(@role, $comment-reference-style-regex)]]
                        [every $n in node() satisfies $n/self::w:commentReference]" 
                mode="wml-to-dbk" priority="3">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:key name="by-w14-paraId" match="*[@w14:paraId]" use="@w14:paraId">
    <!-- found on w:p -->
  </xsl:key>
  
  <xsl:key name="by-w15-paraId" match="*[@w15:paraId]" use="@w15:paraId">
    <!-- found on w15:commentEx in commentsExtended.xml -->
  </xsl:key>

  <xsl:template match="w:comment" mode="comment">
    <xsl:param name="comment-id" as="xs:string" tunnel="yes"/>
    <xsl:variable name="para-ids" as="xs:string*" select="descendant::*/@w14:paraId"/>
    <xsl:variable name="extended-info" as="element(w15:commentEx)*" 
      select="key('by-w15-paraId', $para-ids)/self::w15:commentEx"/>
    <xsl:variable name="parent-para-ids" as="xs:string*" select="$extended-info/@w15:paraIdParent"/>
    <xsl:variable name="parent-comment-ids" as="xs:string*" 
      select="key('by-w14-paraId', $parent-para-ids)/ancestor::w:comment[1]/@w:id"/>
    <xsl:variable name="done" as="xs:string*" select="distinct-values($extended-info/@w15:done)"/>
    <annotation annotates="comment_{$comment-id}">
      <xsl:if test="exists(@w:author | @w:date | @w:initials) or exists($done) or exists($parent-comment-ids)">
        <info>
          <xsl:if test="@w:author | @w:initials">
            <author>
              <personname>
                <xsl:apply-templates select="@w:author | @w:initials" mode="#current"/>
              </personname>
            </author>
          </xsl:if>
          <xsl:apply-templates select="@w:date" mode="#current"/>
          <xsl:if test="exists($done) or exists($parent-comment-ids)">
            <keywordset>
              <xsl:if test="$done">
                <keyword role="done">
                  <xsl:value-of select="$done" separator=" "/>
                </keyword>
              </xsl:if>
              <xsl:if test="$parent-comment-ids">
                <keyword role="parent-comment-ids">
                  <!-- should be at most one parent comment ID, but this is not guaranteed because the metadata
                    are attached to comment paras, not to a comment as a whole -->
                  <xsl:value-of select="for $p in $parent-comment-ids return concat('comment_', $p)" separator=" "/>
                </keyword>
              </xsl:if>
            </keywordset>
          </xsl:if>
        </info>
      </xsl:if>
      <xsl:apply-templates select="*" mode="wml-to-dbk"/>
    </annotation>
  </xsl:template>

  <xsl:template match="@w:author" mode="comment">
    <othername role="display-name">
      <xsl:value-of select="."/>
    </othername>
  </xsl:template>

  <xsl:template match="@w:initials" mode="comment">
    <othername role="initials">
      <xsl:value-of select="."/>
    </othername>
  </xsl:template>

  <xsl:template match="@w:date" mode="comment">
    <date>
      <xsl:value-of select="."/>
    </date>
  </xsl:template>
  
  <xsl:template match="w:annotationRef" mode="wml-to-dbk"/>

</xsl:stylesheet>