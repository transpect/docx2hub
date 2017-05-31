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
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0"
  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn tr mml docx2hub">

  <!--<xsl:import href="http://transpect.io/xslt-util/hex/xsl/hex.xsl"/>-->

  <xsl:template match="dbk:para[
                         dbk:br[@role eq 'column'][preceding-sibling::node() and following-sibling::node()]
                       ]" mode="docx2hub:join-runs" priority="5">
    <xsl:variable name="context" select="." as="element(dbk:para)"/>
    <xsl:variable name="splitted" as="element(dbk:para)+">
      <xsl:for-each-group select="node()" group-starting-with="dbk:br[@role eq 'column']">
        <para>
          <xsl:sequence select="$context/@*"/>
          <xsl:if test="$context/@srcpath and position() != 1 and $context/@srcpath">
            <xsl:attribute name="srcpath" select="concat($context/@srcpath, ';n=', position())"/>
          </xsl:if>
          <xsl:sequence select="current-group()[not(self::dbk:br[@role eq 'column'])]"/>
        </para>
      </xsl:for-each-group>
    </xsl:variable>
    <xsl:apply-templates select="$splitted" mode="#current"/>
  </xsl:template>

  <!-- w:r is here for historic reasons. We used to group the text runs
       prematurely until we found out that it is better to group when
       there's docbook markup. So we implemented the special case of
       dbk:anchors (corresponds to w:bookmarkStart/End) only for dbk:anchor. 
       dbk:anchors between identically formatted phrases will be merged into
       with the phrases' content into a consolidated phrase. -->
  <xsl:template match="*[w:r or dbk:phrase or dbk:superscript or dbk:subscript]" mode="docx2hub:join-runs" priority="3">
    <!-- move sidebars to para level --><xsl:variable name="context" select="."/>
    <xsl:if test="self::dbk:para and .//dbk:sidebar">
      <xsl:call-template name="docx2hub_move-invalid-sidebar-elements"/>
    </xsl:if>
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:call-template name="docx2hub_pagebreak-elements-to-attributes"/>
      <xsl:for-each-group select="node()" group-adjacent="tr:signature(.)">
        <xsl:choose>
          <xsl:when test="current-grouping-key() eq ''">
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy copy-namespaces="no">
              <xsl:apply-templates select="@role, @* except (@srcpath, @role)" mode="#current"/>
              <xsl:if test="$srcpaths = 'yes' and current-group()/@srcpath">
                <xsl:attribute name="srcpath" select="current-group()/@srcpath" separator=" "/>
              </xsl:if>
              <xsl:apply-templates select="current-group()[not(self::dbk:anchor)]/node() 
                                           union current-group()[self::dbk:anchor]" mode="#current" />
            </xsl:copy>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="*" mode="docx2hub:join-runs">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@role, @* except @role, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <!-- collateral: remove links that don’t link -->
  <xsl:template match="dbk:link[@*][every $a in @* satisfies ('srcpath' = $a/name())]" mode="docx2hub:join-runs" priority="5">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:key name="docx2hub:linking-item-by-id" match="*[@linkend | @linkends]" use="@linkend, tokenize(@linkends, '\s+')"/>
  <xsl:key name="docx2hub:item-by-id" match="*[@xml:id]" use="@xml:id"/>

  <!-- collateral: deflate an adjacent start/end anchor pair to a single anchor --> 
  <xsl:template match="dbk:anchor[
                         following-sibling::node()[1] is key('docx2hub:linking-item-by-id', @xml:id)/self::dbk:anchor
                       ]" mode="docx2hub:join-runs">
    <xsl:copy>
      <xsl:apply-templates select="@* except @role" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="dbk:anchor[
                         preceding-sibling::node()[1] is key('docx2hub:item-by-id', @linkend)
                       ]" mode="docx2hub:join-runs"/>

  <xsl:template match="dbk:anchor/@linkend" mode="docx2hub:join-runs">
    <!-- I’d like to keep it for bookmark ranges, but it isn’t allowed in DocBook -->
  </xsl:template>

  <!-- collateral: create indexterms for what was a bookmark range in docx -->
  <xsl:template match="dbk:anchor[exists(key('docx2hub:linking-item-by-id', @xml:id)/self::dbk:indexterm[@linkends])]"
    mode="docx2hub:join-runs" priority="1">
    <xsl:variable name="next-match" as="element(dbk:anchor)?">
      <xsl:next-match/>  
    </xsl:variable>
    <xsl:sequence select="$next-match"/>
    <xsl:variable name="indexterms" as="element(dbk:indexterm)+" 
      select="key('docx2hub:linking-item-by-id', @xml:id)/self::dbk:indexterm"/>
    <xsl:variable name="context" select="." as="element(dbk:anchor)"/>
    <xsl:for-each select="$indexterms">
      <xsl:variable name="pos" as="xs:integer" select="index-of(tokenize(@linkends, '\s+'), $context/@xml:id)"/>
      <xsl:variable name="id" select="concat('itr_', generate-id())" as="xs:string"/>
      <xsl:choose>
        <xsl:when test="$pos = 1">
          <xsl:copy>
            <xsl:apply-templates select="@* except @linkends" mode="#current"/>
            <xsl:if test="$next-match/@role = ('start', 'hub:start')">
              <xsl:attribute name="xml:id" select="$id"/>
              <xsl:attribute name="class" select="'startofrange'"/>
            </xsl:if>
            <xsl:apply-templates mode="#current"/>
          </xsl:copy>
        </xsl:when>
        <xsl:when test="$pos = 2 and exists($next-match)">
          <xsl:copy>
            <xsl:apply-templates select="@* except @linkends" mode="#current"/>
            <xsl:attribute name="startref" select="$id"/>
            <xsl:attribute name="class" select="'endofrange'"/>
          </xsl:copy>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="dbk:indexterm[@linkends]" mode="docx2hub:join-runs"/>

  <!-- collateral: replace name of mapped symbols with default Unicode font name -->
  <xsl:template match="@css:font-family[. = $docx2hub:symbol-font-names][.. = docx2hub:font-map(.)/symbols/symbol/@char]"
    mode="docx2hub:join-runs">
    <xsl:variable name="target-font" as="xs:string?" select="docx2hub:font-map(.)/symbols/symbol[@char = current()/..][1]/@font"/>
    <xsl:attribute name="{name()}" select="if ($target-font) then $target-font else $docx2hub:symbol-replacement-rfonts/@w:ascii"/>
  </xsl:template>
  
  <xsl:template match="@css:font-family" mode="docx2hub:join-runs" priority="2">
    <xsl:variable name="transformed" as="attribute(css:font-family)">
      <xsl:next-match/>
    </xsl:variable>
    <xsl:variable name="role" as="attribute(role)?">
      <xsl:apply-templates select="../@role" mode="#current"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$transformed = . and ../@docx2hub:map-from = ."><!-- no mapping took place, although it should have -->
        <xsl:attribute name="role" select="string-join(distinct-values((tokenize($role, '\s+'), 'hub:ooxml-symbol')), ' ')"/>
        <xsl:attribute name="annotations" 
          select="string-join(for $i in string-to-codepoints(..) return tr:dec-to-hex($i), ' ')"/>
      </xsl:when>
      <xsl:when test="$role">
        <xsl:attribute name="role" select="$role"/>
      </xsl:when>
    </xsl:choose>
    <xsl:sequence select="$transformed"/>
  </xsl:template>
  
  <xsl:template match="@docx2hub:map-from | @docx2hub:field-function" mode="docx2hub:join-runs"/>


  <xsl:template match="dbk:para" mode="docx2hub:join-runs">
    <xsl:call-template name="docx2hub_move-invalid-sidebar-elements"/>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:call-template name="docx2hub_pagebreak-elements-to-attributes"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template name="docx2hub_move-invalid-sidebar-elements">
    <xsl:for-each select=".//dbk:sidebar">
      <xsl:copy>
        <xsl:apply-templates select="@*" mode="#current"/>
        <xsl:attribute name="linkend" select="concat('id_', generate-id(.))"/>
        <xsl:apply-templates select="node()" mode="#current"/>
      </xsl:copy>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="docx2hub_pagebreak-elements-to-attributes">
    <xsl:apply-templates select=".//dbk:br[@role[not(. eq 'textWrapping')]]
                                          [dbk:same-scope(., current())]" 
      mode="docx2hub:join-runs-br-attr"/>
  </xsl:template>

  <xsl:function name="tr:signature" as="xs:string?">
    <xsl:param name="node" as="node()?" />
    <xsl:variable name="result-strings" as="xs:string*">
      <xsl:apply-templates select="$node" mode="docx2hub:join-runs-signature" />
    </xsl:variable>
    <xsl:value-of select="string-join($result-strings,'')"/>
  </xsl:function>

  <xsl:template match="dbk:phrase|dbk:superscript|dbk:subscript" mode="docx2hub:join-runs-signature">
    <xsl:sequence select="string-join(
                            (: don't join runs that contain field chars or instrText :)
                            (name(), w:fldChar/@w:fldCharType, w:instrText/name(), tr:attr-hashes(.)), 
                            '___'
                          )" />
  </xsl:template>

  <xsl:template match="dbk:anchor[
                         tr:signature(following-sibling::node()[not(self::dbk:anchor)][1]/self::element())
                         =
                         tr:signature(preceding-sibling::node()[not(self::dbk:anchor)][1]/self::element())
                       ]" mode="docx2hub:join-runs-signature">
    <xsl:apply-templates select="preceding-sibling::node()[not(self::dbk:anchor)][1]" mode="docx2hub:join-runs-signature" />
  </xsl:template>

  <xsl:template match="node()" mode="docx2hub:join-runs-signature">
    <xsl:sequence select="''" />
  </xsl:template>

  <xsl:function name="tr:attr-hashes" as="xs:string*">
    <xsl:param name="elt" as="node()*" />
    <xsl:perform-sort>
      <xsl:sort/>
      <xsl:sequence select="for $a in ($elt/@* except ($elt/@tr:processed, $elt/@srcpath, $elt/@docx2hub:map-from)) 
                            return tr:attr-hash($a)" />
    </xsl:perform-sort>
    <!-- unmappable chars should stay in their own span: --> 
    <xsl:sequence select="$elt[@docx2hub:map-from[. = $elt/@css:font-family]]/generate-id()"/>
  </xsl:function>

  <xsl:function name="tr:attr-hash" as="xs:string">
    <xsl:param name="att" as="attribute(*)" />
    <xsl:sequence select="concat(name($att), '__=__', $att)" />
  </xsl:function>

  <xsl:function name="tr:attname" as="xs:string">
    <xsl:param name="hash" as="xs:string" />
    <xsl:value-of select="replace($hash, '__=__.+$', '')" />
  </xsl:function>

  <xsl:function name="tr:attval" as="xs:string">
    <xsl:param name="hash" as="xs:string" />
    <xsl:value-of select="replace($hash, '^.+__=__', '')" />
  </xsl:function>
  
  <!-- @type = ('column', 'page') --> 
  <xsl:template match="dbk:br[@role[not(. eq 'textWrapping')]]" mode="docx2hub:join-runs-br-attr">
    <xsl:choose>
      <xsl:when test="dbk:before-text-in-para(., ancestor::dbk:para[1])">
        <xsl:attribute name="css:page-break-before" select="'always'"/>
      </xsl:when>
      <xsl:when test="dbk:after-text-in-para(., ancestor::dbk:para[1])">
        <xsl:attribute name="css:page-break-after" select="'always'"/>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>

  <xsl:variable name="dbk:scope-establishing-elements" as="xs:string*"
    select="('annotation', 
             'entry', 
             'blockquote', 
             'figure', 
             'footnote',
             'indexterm',
             'listitem',
             'sidebar',
             'table')"/>

  <xsl:function name="dbk:same-scope" as="xs:boolean">
    <!-- 
      There are situations when you don’t want to select the
      text nodes of an embedded footnote when selecting the text
      nodes of a paragraph.
      A footnote, for example, constitutes a so called “scope.”
      Other scope-establishing elements are table cells that
      may contain paragraphs, or figures/tables whose captions 
      may contain paragraphs. But also indexterms, elements that 
      do not contain paragraphs, may establish a new scope. 
      This concept allows you to select only the main narrative 
      text of a given paragraph (or phrase, …), excluding any 
      content of embedded notes, figures, list items, or index 
      terms.
      Example:
<para><emphasis>Outer</emphasis> para text<footnote><para>Footnote text</para></footnote>.</para>
      Typical invocation (context: outer para):
      .//text()[ens:same-scope(., current())]
      Result: The three text nodes with string content
      'Outer', ' para text', and '.'
      -->
    <xsl:param name="node" as="node()" />
    <xsl:param name="ancestor-elt" as="element(*)*" />
    <xsl:sequence 
      select="not(
                $node/ancestor::*[
                  local-name() = $dbk:scope-establishing-elements]
                  [
                    some $a in ancestor::* 
                    satisfies (
                      some $b in $ancestor-elt 
                      satisfies ($a is $b))
                  ]
                )" />
  </xsl:function>
  
  <xsl:function name="dbk:before-text-in-para" as="xs:boolean">
    <xsl:param name="elt" as="element(*)"/><!-- typically dbk:br[@role = 'page'] -->
    <xsl:param name="para" as="element(dbk:para)"/>
    <xsl:sequence select="not($para//text())
                          or
                          (
                            dbk:same-scope($elt, $para)
                            and
                            not( some $text in $para//text()[dbk:same-scope(., $para)] 
                                 satisfies ($text &lt;&lt; $elt) 
                            )
                          )"/>
  </xsl:function>

  <xsl:function name="dbk:after-text-in-para" as="xs:boolean">
    <xsl:param name="elt" as="element(*)"/><!-- typically dbk:br[@role = 'page'] -->
    <xsl:param name="para" as="element(dbk:para)"/>
    <xsl:sequence select="dbk:same-scope($elt, $para)
                          and
                          not( some $text in $para//text()[dbk:same-scope(., $para)] 
                               satisfies ($text &gt;&gt; $elt) 
                          )"/>
  </xsl:function>

  <xsl:template match="dbk:br[@role[not(. eq 'textWrapping')]]
                             [
                               dbk:before-text-in-para(., ancestor::dbk:para[1])
                               or dbk:after-text-in-para(., ancestor::dbk:para[1])
                             ]" mode="docx2hub:join-runs"/>

  <!-- sidebar -->
  <xsl:template match="dbk:sidebar" mode="docx2hub:join-runs">
    <anchor>
      <xsl:attribute name="xml:id" select="concat('side_', generate-id(.))"/>
    </anchor>
  </xsl:template>

  <!-- mode hub:fix-libre-office-issues -->
  <!-- style names from Libre Office templates are coded in @native-name only, the @name attributes are automatically numbered 'style65' etc.
        therefore the @names are replace by normalized @native-names in paragraphs and css:rules -->
  
  <xsl:function name="docx2hub:normalize-to-css-name" as="xs:string">
    <xsl:param name="style-name" as="xs:string"/>
    <xsl:sequence select="replace(replace(replace($style-name, '[^-_~a-z0-9]', '_', 'i'), '~', '_-_'), '^(\I)', '_$1')"/>
<!--    <xsl:sequence select="replace($style-name, '~', '_-_')"/>-->
  </xsl:function>
  
  <xsl:key name="natives" match="css:rule" use="@name"/> 
  
  <xsl:variable name="is-libre-office-document"
              select="if (/dbk:hub/dbk:info/dbk:keywordset/dbk:keyword[@role = 'source-application'][matches(., '^LibreOffice', 'i')]) 
                      then true() 
                      else false()" as="xs:boolean"/> 
  
  <xsl:template match="css:rule[$is-libre-office-document][matches(@native-name, '^p$')]/@native-name" mode="docx2hub:join-runs">
    <xsl:attribute name="{name()}">
      <xsl:sequence select="'para'"/>
    </xsl:attribute>
  </xsl:template>
  
  <!-- matches(@native-name, '~'): use @native-name instead of @name even if docx is saved by MS Word -->
  <xsl:template match="css:rule[$is-libre-office-document or matches(@native-name, '~')]/@name" mode="docx2hub:join-runs">
    <xsl:attribute name="{name()}">
      <xsl:choose>
        <xsl:when test="matches(../@native-name, '(Kein Absatzformat|^\s*$)')">
          <xsl:sequence select="'None'"/>
        </xsl:when>
        <xsl:when test="matches(../@native-name, '^p$')">
          <xsl:sequence select="'para'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="docx2hub:normalize-to-css-name(../@native-name)"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute> 
  </xsl:template>
  
  <xsl:template match="*[not((local-name(.) = ('keyword', 'keywordset', 'anchor')))]
                        [$is-libre-office-document or matches(key('natives', @role)/@native-name, '~')]/@role" mode="docx2hub:join-runs">
    <xsl:attribute name="{name()}">
      <xsl:choose>
        <xsl:when test="key('natives', .)[matches(@native-name, 'Kein Absatzformat')]">
          <xsl:sequence select="'None'"/>
        </xsl:when>
        <xsl:when test="key('natives', .)[matches(@native-name, '(Einfaches Absatzformat|^p$)', 'i')]">
          <xsl:sequence select="'para'"/>
        </xsl:when>
        <xsl:when test="matches(., 'hub:')">
          <xsl:sequence select="."/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="key('natives', .)/docx2hub:normalize-to-css-name(@native-name)"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute> 
  </xsl:template>
  
  <xsl:template match="*:keywordset[@role='fieldVars']" mode="docx2hub:join-runs">
    <xsl:if test="exists(//*:keyword[matches(@role,'^fieldVar_')])">
      <xsl:copy>
        <xsl:apply-templates select="@*" mode="#current"/>
        <xsl:apply-templates select="//*:keyword[matches(@role,'^fieldVar_')]" mode="field-var"/>
      </xsl:copy>  
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="*:keyword[matches(@role,'^fieldVar_')]" mode="field-var">
    <xsl:copy>
      <xsl:attribute name="role" select="replace(@role,'^fieldVar_','')"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:keyword[matches(@role,'^fieldVar_')]" mode="docx2hub:join-runs"/>
  
  <!-- This mode has to run before docx2hub:field-functions --> 
  <xsl:template match="*[w:r/w:instrText]" mode="docx2hub:join-instrText-runs">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="node()" 
        group-adjacent="exists(self::w:r[w:instrText][every $c in * satisfies $c/(self::w:instrText | self::w:br (: w:br appeared in comments in 12181_2015_0024_Manuscript.docm :))]
                               | self::w:bookmarkStart | self::w:bookmarkEnd)"><!-- the _GoBack bookmark might be here -->
        <xsl:choose>
          <xsl:when test="current-grouping-key() and exists(current-group()/self::w:r)">
            <w:r>
              <xsl:apply-templates select="current-group()/self::w:r/(@* except @srcpath)" mode="#current"/>
              <xsl:if test="$srcpaths = 'yes' and current-group()/@srcpath">
                <xsl:attribute name="srcpath" select="current-group()/@srcpath" separator=" "/>
              </xsl:if>
              <w:instrText xsl:exclude-result-prefixes="#all">
                <xsl:sequence select="string-join(current-group()/w:instrText, '')"/>
              </w:instrText>
            </w:r>
            <xsl:sequence select="current-group()/(self::w:bookmarkStart | self::w:bookmarkEnd)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="w:footnote/w:p[1][*[docx2hub:element-is-footnoteref(.)]]" mode="docx2hub:join-instrText-runs" priority="1">
    <xsl:variable name="prelim" as="document-node(element(*))">
      <xsl:document>
        <xsl:call-template name="docx2hub:first-note-para"/>
      </xsl:document>
    </xsl:variable>
    <!-- the current template won’t match on $prelim, therefore the template above, that matches *[w:r/w:instrText], 
      should match if there is w:instrText -->
    <xsl:apply-templates select="$prelim" mode="#current"/>
  </xsl:template>

  <!-- Making field function nesting explicit:
       – 'separate' and 'end' fldChars link to their 'begin' fldChar
       – the nesting level is given at each 'begin' fldChar, starting from 1 for the topmost nesting
   -->
  <xsl:template match="/" mode="docx2hub:join-instrText-runs">
    <xsl:variable name="nested-field-functions" as="document-node()">
      <xsl:call-template name="docx2hub:nest-field-functions">
        <xsl:with-param name="input">
          <xsl:document>
            <xsl:sequence select="//w:fldChar"/>
          </xsl:document>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:next-match>
      <xsl:with-param name="nested-field-functions" select="$nested-field-functions" tunnel="yes"/>
    </xsl:next-match>
  </xsl:template>
  
  <!-- this happened in a previous mode: -->
  <xsl:template match="w:fldChar" mode="docx2hub:remove-redundant-run-atts">
    <xsl:copy>
      <xsl:attribute name="xml:id" select="concat('fldChar_', generate-id())"/>
      <xsl:apply-templates select="@*" mode="#current"/>
    </xsl:copy>    
  </xsl:template>
  
  <!-- also in a previous mode, remove w:lastRenderedPageBreak as it may be in
  runs that otherwise contain w:instrText, preventing adjacent w:instrText runs from being merged -->
  <xsl:template match="w:lastRenderedPageBreak" mode="docx2hub:remove-redundant-run-atts"/>
  
  <xsl:template name="docx2hub:nest-field-functions" as="document-node()">
    <xsl:param name="input" as="document-node()"/><!-- containing raw w:fldChar or nested docx2hub:field-function -->
    <xsl:variable name="grouping" as="element(*)*">
      <xsl:for-each-group select="$input/*" 
                          group-starting-with="w:fldChar[@w:fldCharType = 'begin']
                                                        [following-sibling::w:fldChar[1]/@w:fldCharType = 'separate']
                                                        [following-sibling::w:fldChar[2]/@w:fldCharType = 'end']
                                               | 
                                               w:fldChar[@w:fldCharType = 'begin']
                                                        [following-sibling::w:fldChar[1]/@w:fldCharType = 'end']">
        <xsl:choose>
          <xsl:when test="self::w:fldChar[@w:fldCharType = 'begin']
                                         [following-sibling::w:fldChar[1]/@w:fldCharType = 'separate']
                                         [following-sibling::w:fldChar[2]/@w:fldCharType = 'end']
                          |
                          self::w:fldChar[@w:fldCharType = 'begin']
                                         [following-sibling::w:fldChar[1]/@w:fldCharType = 'end']">
            <xsl:variable name="end" as="element(w:fldChar)" 
              select="self::w:fldChar[@w:fldCharType = 'begin']
                              /following-sibling::w:fldChar[position() = (1, 2)][@w:fldCharType = 'end'][1]"/>
            <docx2hub:field-function>
              <xsl:apply-templates select="current-group()[. &lt;&lt; $end] union $end" mode="#current">
                <xsl:with-param name="begin" as="element(w:fldChar)" select="." tunnel="yes"/>
              </xsl:apply-templates>
            </docx2hub:field-function>
            <xsl:apply-templates select="current-group()[. &gt;&gt; $end]" mode="#current"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="count($grouping) = count($input/*)">
        <xsl:document>
          <xsl:sequence select="$grouping"/>
        </xsl:document>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="docx2hub:nest-field-functions">
          <xsl:with-param name="input">
            <xsl:document>
              <xsl:sequence select="$grouping"/>
            </xsl:document>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:fldChar[@w:fldCharType = ('separate', 'end')]/@xml:id" mode="docx2hub:join-instrText-runs">
    <!-- $begin is present while the field functions are nested, detached from their original context: --> 
    <xsl:param name="begin" as="element(w:fldChar)?" tunnel="yes"/>
    <!-- $nested-field-functions is present when the complete document is processed (after xsl:next-match for /),
      adding the nesting info (@linkend that points to the begin fldChar) to the separate/end fldChars: --> 
    <xsl:param name="nested-field-functions" as="document-node()?" tunnel="yes"/>
    <xsl:copy/>
    <xsl:if test="exists($begin)">
      <xsl:attribute name="linkend" select="$begin/@xml:id"/>
    </xsl:if>
    <xsl:if test="exists($nested-field-functions)">
      <xsl:variable name="corresponding-item-in-nesting" as="element(w:fldChar)" 
        select="key('docx2hub:item-by-id', ., $nested-field-functions)"/>
      <xsl:sequence select="$corresponding-item-in-nesting/@linkend"/>
      <xsl:attribute name="level" select="count($corresponding-item-in-nesting/ancestor::docx2hub:field-function)"/>
    </xsl:if>
  </xsl:template>
  
  <!-- determine whether inline or display equation -->
  
  <xsl:template match="dbk:inlineequation[@role eq 'mtef']" mode="docx2hub:join-runs">
    <xsl:variable name="para" select="ancestor::dbk:para[1]" as="element(dbk:para)"/>
    <xsl:variable name="is-display-equation" 
      select="count($para//dbk:inlineequation[dbk:same-scope(., $para)]) eq 1
              and 
              not(
                normalize-space(
                  string-join(
                    $para//text()[dbk:same-scope(., $para)]
                                 [namespace-uri(..) ne 'http://www.w3.org/1998/Math/MathML'],
                    ''
                  )
                )
              )" as="xs:boolean"/>
    <xsl:element name="{if ($is-display-equation) then 'equation' else name()}">
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:element>
  </xsl:template>
  
<!-- group more than one mml:mi[@mathvariant='normal'] element to mtext; exclude mathtype2mml processed mathml -->
  <xsl:template match="mml:*[not(ancestor::*[ends-with(name(), 'equation')]/@role eq 'mtef')][mml:mi]" mode="docx2hub:join-runs">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="node()" group-adjacent="exists(self::mml:mi[@mathvariant eq 'normal'] | self::mml:mtext)">
        <xsl:choose>
          <xsl:when test="current-grouping-key() and string-length(string-join(current-group(), '')) gt 1">
            <xsl:variable name="prelim" as="element(mml:mtext)">
              <mml:mtext>
                <xsl:apply-templates select="current-group()/node()" mode="#current"/>
              </mml:mtext>  
            </xsl:variable>
            <xsl:apply-templates select="$prelim" mode="#current"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="mml:mtext" mode="docx2hub:join-runs">
    <xsl:call-template name="mtext-or-mspace">
      <xsl:with-param name="string" as="xs:string" select="."/>
    </xsl:call-template>
  </xsl:template>
  
  <xsl:variable name="opening-parenthesis" select="('[', '{', '(')" as="xs:string*"/>
  <xsl:variable name="closing-parenthesis" select="(']', '}', ')')" as="xs:string*"/>
  
  <xsl:template match="mml:mrow[count(*) = 3]
                              [*[1]/self::mml:mo = $opening-parenthesis]
                              [*[3]/self::mml:mo = $closing-parenthesis]" mode="docx2hub:join-runs" priority="1">
    <mml:mfenced open="{*[1]}" close="{*[3]}" separators="">
      <xsl:apply-templates select="if(*[2]/self::mml:mrow) then *[2]/node() else *[2]" mode="#current"/>
    </mml:mfenced>
  </xsl:template>
  
</xsl:stylesheet>
