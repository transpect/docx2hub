<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:extendedProps="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:tr="http://transpect.io"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:w10="urn:schemas-microsoft-com:office:word" 
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
  xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  exclude-result-prefixes="docx2hub mml tr dbk cp"
  version="2.0">
  
  <xsl:param name="srcpaths-on-runs" as="xs:string" select="'yes'"/>
  <!-- insert-default-settings: add default lang and font or only add srcpaths atts in mode insert-xpath -->
  <xsl:param name="insert-default-settings" as="xs:string" select="'yes'"/>

  <xsl:function name="docx2hub:srcpath" as="xs:string">
    <xsl:param name="elt" as="element(*)?"/>
    <xsl:sequence select="string-join(
                            (
                              if ($elt/.. instance of element(*)) then docx2hub:srcpath($elt/..) else concat(base-uri($elt), '?xpath='),
                              '/',
                              name($elt),
                              '[',
                              xs:string(index-of(for $s in $elt/../*[name() = name($elt)] return generate-id($s), generate-id($elt))),
                              ']'
                            ),
                            ''
                          )"/>
  </xsl:function>

  <xsl:variable name="docx2hub:srcpath-elements" as="xs:string+"
    select="if ($srcpaths-on-runs = 'yes')
            then ('w:p',
                  'w:t',
                  'w:tbl',
                  'w:tc',
                  'w:hyperlink',
                  'w:r',
                  'w:tab',
                  'w:br',
                  'v:imagedata',
                  'o:OLEObject',
                  'w:drawing',
                  'w:comment',
                  'w:endnote',
                  'w:footnote'
                 )
             else ('w:p',
                  'w:tbl',
                  'w:tc',
                  'w:hyperlink',
                  'w:tab',
                  'w:br',
                  'v:imagedata',
                  'o:OLEObject',
                  'w:drawing',
                  'w:comment',
                  'w:endnote',
                  'w:footnote'
                 )"/>

  <xsl:template match="*[ $srcpaths = 'yes' ]
                        [ name() = $docx2hub:srcpath-elements ]
                        [ not(ancestor::w:pPr) (: suppress srcpath in tab declarations :)] 
                        [ /*/name() = ('w:document', 'w:footnotes', 'w:endnotes', 'w:settings', 
                                       'w:comments', 'w:header', 'w:footer', 'w:hdr', 'w:ftr')]
                        " mode="insert-xpath">
    <xsl:copy>
      <xsl:attribute name="srcpath" select="docx2hub:srcpath(.)"/>
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/><!-- there was a case where the input was overwritten with an output that contained srcpaths -->
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  
  <xsl:template match="/*" mode="insert-xpath" priority="-0.2">
    <!-- no copy-namespaces="no" in order to suppress excessive namespace declarations on every element -->
    <xsl:copy>
      <xsl:namespace name="w16cid">http://schemas.microsoft.com/office/word/2016/wordml/cid</xsl:namespace>
      <xsl:namespace name="w16se">http://schemas.microsoft.com/office/word/2015/wordml/symex</xsl:namespace>
      <xsl:namespace name="tr">http://transpect.io</xsl:namespace>
      <xsl:namespace name="m">http://schemas.openxmlformats.org/officeDocument/2006/math</xsl:namespace>
      <xsl:namespace name="mml">http://www.w3.org/1998/Math/MathML</xsl:namespace>
      <xsl:namespace name="o">urn:schemas-microsoft-com:office:office</xsl:namespace>
      <xsl:namespace name="pkg">http://schemas.microsoft.com/office/2006/xmlPackage</xsl:namespace>
      <xsl:namespace name="r">http://schemas.openxmlformats.org/officeDocument/2006/relationships</xsl:namespace>
      <xsl:namespace name="rel">http://schemas.openxmlformats.org/package/2006/relationships</xsl:namespace>
      <xsl:namespace name="v">urn:schemas-microsoft-com:vml</xsl:namespace> 
      <xsl:namespace name="w">http://schemas.openxmlformats.org/wordprocessingml/2006/main</xsl:namespace>
      <xsl:namespace name="word200x">http://schemas.microsoft.com/office/word/2003/wordml</xsl:namespace>
      <xsl:namespace name="wx">http://schemas.microsoft.com/office/word/2003/auxHint</xsl:namespace>
      <xsl:namespace name="w14">http://schemas.microsoft.com/office/word/2010/wordml</xsl:namespace>
      <xsl:namespace name="wp">http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing</xsl:namespace>
      <xsl:namespace name="cp">http://schemas.openxmlformats.org/package/2006/metadata/core-properties</xsl:namespace>
      <xsl:namespace name="extendedProps">http://schemas.openxmlformats.org/officeDocument/2006/extended-properties</xsl:namespace>
      <xsl:namespace name="customProps">http://schemas.openxmlformats.org/officeDocument/2006/custom-properties</xsl:namespace>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="/w:document" mode="insert-xpath" as="document-node(element(w:root))" priority="2">
    <xsl:variable name="docRels-uri" as="xs:anyURI"
      select="if (doc-available(resolve-uri(concat($base-dir,'_rels/document2.xml.rels'))))
                  then resolve-uri(concat($base-dir,'_rels/document2.xml.rels'))
                  else resolve-uri(concat($base-dir,'_rels/document.xml.rels'))"/>
    <xsl:document>
      <w:root>
        <xsl:attribute name="xml:base" select="$container-base-uri"/>
        <xsl:attribute name="extract-dir-uri" select="$extract-dir-uri"/>
        <xsl:attribute name="local-href" select="$local-href"/>
        <xsl:variable name="containerRels" as="document-node(element(rel:Relationships))"
          select="document(resolve-uri(replace($base-dir, '[^/]+/?$', '_rels/.rels')))"/>
        <xsl:apply-templates select="document(resolve-uri('styles.xml', $base-dir))/w:styles" mode="#current"/>
        <xsl:if test="doc-available(resolve-uri('numbering.xml', $base-dir))">
          <xsl:apply-templates select="document(resolve-uri('numbering.xml', $base-dir))/w:numbering" mode="#current"/>
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('footnotes.xml', $base-dir))">
          <xsl:apply-templates select="document(resolve-uri('footnotes.xml', $base-dir))/w:footnotes" mode="#current"/>
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('endnotes.xml', $base-dir))">
          <xsl:apply-templates select="document(resolve-uri('endnotes.xml', $base-dir))/w:endnotes" mode="#current"/>
        </xsl:if>
        <xsl:apply-templates select="document(resolve-uri('settings.xml', $base-dir))/w:settings" mode="#current"/>
        <xsl:if test="doc-available(resolve-uri('comments.xml', $base-dir))">
          <xsl:apply-templates select="document(resolve-uri('comments.xml', $base-dir))/w:comments" mode="#current"/>
        </xsl:if>
        <xsl:apply-templates select="document(resolve-uri('fontTable.xml', $base-dir))/w:fonts" mode="#current"/>
        <w:docTypes>
          <xsl:apply-templates select="document(resolve-uri('../%5BContent_Types%5D.xml', $base-dir))/ct:Types" mode="#current"
          />
        </w:docTypes>
        <w:containerRels>
          <xsl:apply-templates select="$containerRels/*" mode="#current"/>
        </w:containerRels>
        <w:containerProps>
          <!-- custom file properties (§ 15.2.12.2) that are found in, e.g., 
            <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
            will later be converted to keywordset[@role = 'hub:custom-meta'] -->
          <xsl:apply-templates mode="#current"
            select="for $t in $containerRels/rel:Relationships/rel:Relationship[ends-with(@Type, 'properties')]/@Target
                    return document(resolve-uri($t, $container-base-uri))"/>
        </w:containerProps>
        <w:docRels>
          <xsl:apply-templates select="document($docRels-uri)/rel:Relationships"
            mode="#current"/>
        </w:docRels>
        <xsl:if test="doc-available(resolve-uri('_rels/footnotes.xml.rels', $base-dir))">
          <w:footnoteRels>
            <xsl:apply-templates select="document(resolve-uri('_rels/footnotes.xml.rels', $base-dir))/rel:Relationships"
              mode="#current"/>
          </w:footnoteRels>
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('_rels/endnotes.xml.rels', $base-dir))">
          <w:endnoteRels>
            <xsl:apply-templates select="document(resolve-uri('_rels/endnotes.xml.rels', $base-dir))/rel:Relationships"
              mode="#current"/>
          </w:endnoteRels>
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('_rels/comments.xml.rels', $base-dir))">
          <w:commentRels>
            <xsl:apply-templates select="document(resolve-uri('_rels/comments.xml.rels', $base-dir))/rel:Relationships"
              mode="#current"/>
          </w:commentRels>
        </xsl:if>

        <!-- apply header and footer content and relations -->
        <xsl:for-each select="('header', 'footer')">
          <xsl:variable name="type" select="current()" as="xs:string"/>
          <xsl:if
            test="some $rel-file
                        in document($docRels-uri)/rel:Relationships/rel:Relationship[
                          @Type eq concat('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', $type)
                        ]/@Target
                        satisfies doc-available(resolve-uri(concat('_rels/', $rel-file, '.rels'), $base-dir))">
            <xsl:element name="w:{$type}Rels">
              <xsl:for-each
                select="document($docRels-uri)/rel:Relationships/rel:Relationship[
                                      @Type eq concat('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', $type)
                                    ]">
                <xsl:if test="doc-available(resolve-uri(concat('_rels/', @Target, '.rels')))">
                  <xsl:apply-templates
                    select="document(resolve-uri(concat('_rels/', @Target, '.rels')))/rel:Relationships"
                    mode="#current"/>
                </xsl:if>
              </xsl:for-each>
            </xsl:element>
          </xsl:if>
          <xsl:element name="w:{$type}">
            <xsl:for-each
              select="document($docRels-uri)/rel:Relationships/rel:Relationship[
                                  @Type eq concat('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', $type)
                                ]">
              <xsl:if test="doc-available(resolve-uri(@Target, $base-dir))">
                <xsl:apply-templates select="document(resolve-uri(@Target, $base-dir))" mode="#current"/>
              </xsl:if>
            </xsl:for-each>
          </xsl:element>
        </xsl:for-each>

        <!-- reproduce the document (with srcpaths), using the default identity template from catch-all.xsl: -->
        <xsl:next-match/>
      </w:root>
    </xsl:document>
  </xsl:template>

  <xsl:template match="  w:document | w:numbering | w:endnotes | w:footnotes
                       | w:settings | w:fonts | rel:Relationships | w:comments
                       | w:styles | extendedProps:Properties
                       | ct:Types | w:hdr | w:ftr | *:Properties | cp:coreProperties" mode="insert-xpath">
    <xsl:copy>
      <xsl:attribute name="xml:base" select="base-uri()" />
      <xsl:apply-templates select="@*, *" mode="#current"/>      
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>
