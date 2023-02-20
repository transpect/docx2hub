<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:w= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  exclude-result-prefixes="dbk docx2hub xs fn xlink w r"
  xmlns="http://docbook.org/ns/docbook"
  version="3.0">

  
  <xsl:template match="CITAVI_JSON" mode="wml-to-dbk tables" 
                use-when="xs:decimal(system-property('xsl:version')) ge 3.0">
    <xsl:param name="citavi-refs" as="document-node()?" tunnel="yes"/>
    <xsl:variable name="citavi-xml" as="element(Placeholder)*"
                  select="for $i in replace(w:bookmarkStart[1]/@w:name, '^_CTVP001', '')
                          return $citavi-refs/docx2hub:citavi-xml/Placeholder[replace(Id, '-', '') eq $i]"/>
    <xsl:choose>
      <xsl:when test="every $n in node()[
                        not(
                          not(normalize-space()) 
                          or
                          self::w:r[
                            every $m in node() 
                            satisfies $m[self::w:fldChar or self::w:instrText]
                          ]
                        )
                      ] satisfies 
                        if($n[self::w:hyperlink[@r:id]])
                        then $n[self::w:hyperlink[@r:id]]
                               /key('docrel', $n/@r:id)/@Target[starts-with(., '#_CTVL')]
                        else //*:CITAVI_XML//w:p[.//w:bookmarkStart/@w:name = $n[self::w:hyperlink]/@w:anchor]">
        <xsl:for-each select="w:hyperlink">
          <xsl:choose>
            <xsl:when test="@r:id">
              <citation docx2hub:citavi-rendered-linkend="{key('docrel', @r:id)/@Target/substring-after(., '#')[not(. = '')]}">
                <xsl:apply-templates select="@srcpath, ancestor::w:sdt[1]/w:sdtPr/w:tag/@w:val" mode="wml-to-dbk">
                  <xsl:with-param name="ref-pos" as="xs:integer" select="position()" tunnel="no"/>
                </xsl:apply-templates>
                <xsl:apply-templates mode="wml-to-dbk"/>
              </citation>
            </xsl:when>
            <xsl:otherwise>
              <citation docx2hub:citavi-rendered-linkend="{@w:anchor}">
                <xsl:apply-templates select="@srcpath, ancestor::w:sdt[1]/w:sdtPr/w:tag/@w:val" mode="wml-to-dbk">
                  <xsl:with-param name="ref-pos" as="xs:integer" select="position()" tunnel="no"/>
                </xsl:apply-templates>
                <xsl:apply-templates mode="wml-to-dbk"/>
              </citation>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="*:HYPERLINK/@fldArgs[matches(., '^.*&#34;#?(_CTVL[^&#34;]+)&#34;.*$')]">
        <xsl:for-each select="*:HYPERLINK[@fldArgs/matches(., '&#34;#?_CTVL[^&#34;]+&#34;')]">
          <xsl:variable name="target" 
            select="replace(@fldArgs, '^.*&#34;#?(_CTVL[^&#34;]+)&#34;.*$', '$1')"/>
          <citation docx2hub:citavi-rendered-linkend="{$target}">
            <xsl:apply-templates select="ancestor::w:sdt[1]/w:sdtPr/w:tag/@w:val" mode="wml-to-dbk">
              <xsl:with-param name="ref-pos" as="xs:integer" select="position()" tunnel="no"/>
            </xsl:apply-templates>
            <xsl:apply-templates mode="wml-to-dbk"/>
          </citation>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="ancestor::w:sdt[1]/w:sdtPr/w:tag/@w:val">
        <citation>
          <xsl:apply-templates select="ancestor::w:sdt[1]/w:sdtPr/w:tag/@w:val" mode="wml-to-dbk"/>
          <xsl:apply-templates mode="wml-to-dbk"/>
        </citation>    
      </xsl:when>
      <xsl:when test="$citavi-xml">
        <citation linkends="{$citavi-xml/Entries/Entry/ReferenceId/concat('_', .)}">
          <xsl:apply-templates mode="wml-to-dbk"/>
        </citation>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates mode="wml-to-dbk"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:sdtPr/w:tag/@w:val[matches(., '^Citavi\.?Placeholder', 'i')]" mode="wml-to-dbk">
    <xsl:param name="citavi-refs" as="document-node()?" tunnel="yes"/>
    <xsl:param name="ref-pos" select="0" as="xs:integer" tunnel="no"/>
    <xsl:if test="exists($citavi-refs/docx2hub:citavi-jsons)">
      <xsl:variable name="cited-refs" as="element(fn:map)*" 
        select="key('docx2hub:by-citavi-placeholder', ., $citavi-refs)
                  /fn:array[@key = 'Entries']
                    /fn:map
                      /fn:map[@key = 'Reference']"/>
      <xsl:attribute name="linkends" separator=" " 
        select="(for $cid in $cited-refs/fn:string[@key = 'Id'] return '_' || $cid)[if($ref-pos != 0) then (position() = $ref-pos) else true()]"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:key name="docx2hub:by-citavi-placeholder" match="fn:map[fn:string[@key = 'Tag']]" 
    use="string(fn:string[@key = 'Tag'])"/>
  
  <xsl:template match="w:sdt[w:sdtPr/w:tag/@w:val[matches(., '^Citavi\.?Placeholder', 'i')]]" mode="wml-to-dbk tables">
    <xsl:apply-templates select="w:sdtContent/*" mode="#current"/>
  </xsl:template>

  <xsl:template name="docx2hub:citavi-json-to-xml" use-when="xs:decimal(system-property('xsl:version')) ge 3.0">
    <xsl:variable name="jsons-actually-containing-xml" as="document-node(element(Placeholder))*" 
                  select="for $jd in .//CITAVI_JSON/@fldArgs 
                          return if(doc-available($jd)) then doc($jd) else ()"/>
    <xsl:variable name="jsons" as="item()*">
      <xsl:for-each select=".//CITAVI_JSON/@fldArgs">
        <xsl:try select="json-to-xml(unparsed-text(current()))">
          <xsl:catch/>
        </xsl:try>
      </xsl:for-each>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="exists($jsons)">
        <xsl:document>
          <docx2hub:citavi-jsons>
            <xsl:sequence select="$jsons"/>
          </docx2hub:citavi-jsons>
        </xsl:document>
      </xsl:when>
      <xsl:when test="exists($jsons-actually-containing-xml)">
        <xsl:document>
          <docx2hub:citavi-xml>
            <xsl:sequence select="$jsons-actually-containing-xml"/>
          </docx2hub:citavi-xml>
        </xsl:document>
      </xsl:when>
    </xsl:choose>
    
  </xsl:template>
  
  <xsl:template match="node() | @*" mode="citavi csl" priority="-1">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template name="citavi-test">
    <xsl:variable name="citavi-bib" as="element(dbk:biblioentry)*">
      <xsl:for-each-group
        select="/docx2hub:citavi-jsons/fn:map/fn:array/fn:map/fn:map[@key = 'Reference']"
        group-by="fn:string[@key = 'Id']">
        <xsl:apply-templates select="." mode="citavi"/>
      </xsl:for-each-group>
    </xsl:variable>
    <xsl:if test="exists($citavi-bib)">
      <bibliography role="Citavi">
        <xsl:sequence select="$citavi-bib"/>
      </bibliography>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="fn:string" mode="citavi">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:template match="fn:map[@key = 'Reference']" mode="citavi">
    <!-- pos var: is there a trustworthy mechanism (i.e. document order)
                  to find the corresponding rendered paragraph? -->
    <xsl:variable name="pos" 
      select="0(:index-of(
                for $i in //fn:map[@key = 'Reference']
                                  [not(../fn:string[@key = 'ReferenceId'] = ../preceding::fn:string[@key = 'ReferenceId'])] 
                  return generate-id($i), 
                generate-id(.)
              ):)" as="xs:integer?"/>
    <biblioentry xml:id="_{fn:string[@key = 'Id']}">
      <xsl:call-template name="citavi-rendered">
        <xsl:with-param name="pos" select="$pos"/>
      </xsl:call-template>
      <xsl:apply-templates mode="#current" 
        select="fn:map[@key = 'ParentReference'][*/@key]"/>
      <xsl:call-template name="citavi-reference"/>
    </biblioentry>
  </xsl:template>

  <xsl:template name="citavi-rendered">
    <xsl:param name="pos" select="0" as="xs:integer"/>
    <xsl:variable name="rendered" 
      select="$root//w:sdt[
                w:sdtPr/w:tag/@w:val = 'CitaviBibliography' and w:sdtContent/*
              ]/w:sdtContent
                /(*:CITAVI_XML/w:p union */self::w:p)[
                  .//w:bookmarkStart[@w:name/starts-with(., '_CTVL')]
                ][position() = $pos]"/>
    <xsl:if test="$rendered">
      <abstract role="rendered">
        <para>
          <xsl:apply-templates select="$rendered/(@srcpath, node())" mode="wml-to-dbk"/>
        </para>
      </abstract>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:sdtContent/*:CITAVI_XML//w:bookmarkEnd" mode="wml-to-dbk"/>
  <xsl:template match="w:sdtContent/*:CITAVI_XML//w:bookmarkStart" mode="wml-to-dbk">
    <anchor xml:id="{@w:name}" role="docx2hub:citavi-rendered"/>
  </xsl:template>
  
  <xsl:template match="Reference" mode="citavi">
    <biblioentry xml:id="_{Id}">
      <xsl:apply-templates select="ParentReference" mode="#current"/>
      <xsl:call-template name="citavi-reference-xml"/>
      <xsl:call-template name="citavi-rendered">
        <xsl:with-param name="pos" select="0(:position():)"/><!-- untested -->
      </xsl:call-template>
    </biblioentry>
  </xsl:template>
  
  <xsl:variable name="other-citavi-ref-parts" as="xs:string+"
                select="'Date',
                        'Doi', 
                        'Edition', 
                        'Isbn', 
                        'Language',
                        'Number',
                        'PageCount',
                        'PageRange',
                        'PlaceOfPublication',
                        'Price',
                        'ShortTitle',
                        'Subtitle',
                        'SourceOfBibliographicInformation',
                        'Title',
                        'Volume',
                        'Year'"/>
  
  <xsl:template name="citavi-reference">
    <xsl:variable name="reference-type" as="attribute(relation)?">
      <xsl:apply-templates select="fn:string[@key = 'ReferenceType']" mode="#current"/>
    </xsl:variable>
    <xsl:apply-templates mode="#current"
                         select="fn:map[@key = 'Periodical']
                                       [fn:string[@key = ('Name', 
                                                          'StandardAbbreviation', 
                                                          'Issn')]]"/>
    <biblioset>
      <xsl:sequence select="$reference-type"/>
      <xsl:apply-templates select="fn:string[@key = ('LanguageCode', 
                                                     'OnlineAddress')]" mode="#current"/>
      <xsl:variable name="authorgroup" as="node()*">
        <xsl:apply-templates select="fn:array[@key = ('Authors', 
                                                      'Editors', 
                                                      'Collaborators')]/fn:map" mode="#current"/>
      </xsl:variable>
      <xsl:if test="exists($authorgroup)">
        <authorgroup>
          <xsl:sequence select="$authorgroup"/>
        </authorgroup>
      </xsl:if>
      <xsl:variable name="publisher" as="element(*)*">
        <xsl:apply-templates select="fn:array[@key = ('Publishers')]/fn:map" mode="#current"/>
      </xsl:variable>
      <xsl:if test="exists($publisher)">
        <publisher>
          <xsl:sequence select="$publisher"/>
        </publisher>
      </xsl:if>
      <xsl:apply-templates select="fn:map[@key = 'SeriesTitle'],
                                   fn:string[@key = $other-citavi-ref-parts]" mode="#current"/>
    </biblioset>
  </xsl:template>
  
  <xsl:template name="citavi-reference-xml">
    <xsl:variable name="reference-type" as="attribute(relation)?">
      <xsl:apply-templates select="ReferenceTypeId" mode="#current"/>
    </xsl:variable>
    <xsl:apply-templates select="Periodical[Issn|Name|StandardAbbreviation]" mode="#current"/>
    <biblioset>
      <xsl:sequence select="$reference-type"/>
      <xsl:apply-templates select="LanguageCode
                                  |OnlineAddress" mode="#current"/>
      <xsl:variable name="authorgroup" as="node()*">
        <xsl:apply-templates select="Authors/Person
                                    |Collaborators/Person
                                    |Editors/Person" mode="#current"/>
      </xsl:variable>
      <xsl:if test="exists($authorgroup)">
        <authorgroup>
          <xsl:sequence select="$authorgroup"/>
        </authorgroup>
      </xsl:if>
      <xsl:variable name="publisher" as="element(*)*">
        <xsl:apply-templates select="Publishers" mode="#current"/>
      </xsl:variable>
      <xsl:if test="exists($publisher)">
        <publisher>
          <xsl:sequence select="$publisher"/>
        </publisher>
      </xsl:if>
      <xsl:apply-templates select="*[local-name() = $other-citavi-ref-parts]" mode="#current"/>
    </biblioset>
  </xsl:template>
  
  <xsl:key name="docx2hub:by-citavi-id" 
           match="fn:map[@key][fn:string[@key = '$id']]
                 |fn:array[@key]/fn:map[empty(@key)][fn:string[@key = '$id']]" 
           use="string-join(((@key, 'person')[1], fn:string[@key = '$id']), '__')"/>
  
  <xsl:template match="fn:map[@key = 'ParentReference'][count(*) = 1][fn:string[@key = '$ref']]" 
                mode="citavi" priority="1">
    <xsl:if test="$debug = 'yes'">
      <xsl:comment>redirect to <xsl:value-of select="@key, fn:string[@key = '$ref']"/></xsl:comment>
    </xsl:if>
    <xsl:message>redirect to <xsl:value-of select="@key, fn:string[@key = '$ref']"/></xsl:message>
    <xsl:call-template name="docx2hub:citavi-redirect"/>
  </xsl:template>
  
  <xsl:template match="fn:map[@key = 'ParentReference']" mode="citavi">
    <xsl:if test="$debug = 'yes'">
      <xsl:comment select="@key, fn:string[@key = '$id']"/>
    </xsl:if>
    <xsl:call-template name="citavi-reference"/>
  </xsl:template>
  
  <xsl:template match="fn:map[@key = 'Periodical']
                      |Periodical" mode="citavi">
    <biblioset relation="journal">
      <xsl:apply-templates select="fn:string[@key = ('Name', 'StandardAbbreviation', 'Issn', 'UserAbbreviation1')]
                                  |Name
                                  |StandardAbbreviation
                                  |Issn
                                  |UserAbbreviation1"
                           mode="#current"/>
    </biblioset>
  </xsl:template>
  
  <xsl:template match="fn:map[@key = 'Periodical']/fn:string[@key = 'Name']
                      |Periodical/Name" 
                mode="citavi">
    <title>
      <xsl:value-of select="."/>
    </title>
  </xsl:template>
  
  <xsl:template match="fn:map[@key = 'Periodical']/fn:string[@key = 'StandardAbbreviation']
                      |Periodical/StandardAbbreviation" 
                mode="citavi">
    <abbrev>
      <xsl:value-of select="."/>
    </abbrev>
  </xsl:template>
  
  <xsl:template match="fn:map[@key = 'Periodical']/fn:string[@key = 'UserAbbreviation1']
                      |Periodical/UserAbbreviation1" 
                mode="citavi">
    <abbrev role="user-abbrev">
      <xsl:value-of select="."/>
    </abbrev>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'ReferenceType']
                      |ReferenceTypeId"
                mode="citavi" as="attribute(relation)?">
    <xsl:choose>
      <xsl:when test=". = 'Book'">
        <xsl:attribute name="relation" select="'book'"/>
      </xsl:when>
      <xsl:when test=". = 'BookEdited'">
        <xsl:attribute name="relation" select="'incollection'"/>
      </xsl:when>
      <xsl:when test=". = 'Broadcast'">
        <xsl:attribute name="relation" select="'broadcast'"/>
      </xsl:when>
      <xsl:when test=". = 'ConferenceProceedings'">
        <xsl:attribute name="relation" select="'proceedings'"/>
      </xsl:when>
      <xsl:when test=". = 'Contribution'">
        <xsl:attribute name="relation" select="'inproceedings'"/>
      </xsl:when>
      <xsl:when test=". = 'InternetDocument'">
        <xsl:attribute name="relation" select="'misc'"/>
      </xsl:when>
      <xsl:when test=". = 'JournalArticle'">
        <xsl:attribute name="relation" select="'article'"/>
      </xsl:when>
      <xsl:when test=". = 'Patent'">
        <xsl:attribute name="relation" select="'patent'"/>
      </xsl:when>
      <xsl:when test=". = 'Unknown'">
        <xsl:attribute name="relation" select="'misc'"/>
      </xsl:when>
      <xsl:when test=". = 'UnpublishedWork'">
        <xsl:attribute name="relation" select="'misc'"/>
      </xsl:when>
      <xsl:when test=". = 'Thesis'">
        <xsl:attribute name="relation" select="'thesis'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
    
  <xsl:template match="fn:string[@key = 'LanguageCode']
                      |LanguageCode" 
                mode="citavi" as="attribute(xml:lang)">
    <xsl:attribute name="xml:lang" select="string(.)"/>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'Language']
                      |Language" 
                mode="citavi">
    <language xmlns="http://purl.org/dc/terms/">
      <xsl:value-of select="."/>
    </language>
    <!--<bibliomisc role="language">
      <xsl:value-of select="."/>
    </bibliomisc>-->
  </xsl:template>
  
  <xsl:template match="fn:array[@key = ('Authors', 'Editors', 'Collaborators')]/fn:map[fn:string[@key = '$ref']]" 
    mode="citavi" priority="1">
    <xsl:if test="$debug = 'yes'">
      <xsl:comment>redirect to <xsl:value-of select="../@key, fn:string[@key = '$ref']"/></xsl:comment>
    </xsl:if>
    <xsl:call-template name="docx2hub:citavi-redirect">
      <xsl:with-param name="id-family" as="xs:string" select="'person'"/>
    </xsl:call-template>
  </xsl:template> 

  <xsl:template name="docx2hub:citavi-redirect">
    <xsl:param name="id-family" select="string((@key, ../@key)[1])" as="xs:string"/>
    <xsl:variable name="same-key" as="element(fn:map)*"
      select="key('docx2hub:by-citavi-id', string-join(($id-family, fn:string[@key = '$ref']), '__'))"/>
    <xsl:variable name="last" select="$same-key[. &lt;&lt; current()][last()]" as="element(fn:map)?"/>
    <xsl:apply-templates select="$last" mode="#current">
      <xsl:with-param name="person-group-name" as="xs:string?" select="parent::fn:array/@key"/>
    </xsl:apply-templates>
  </xsl:template>

  <xsl:template match="fn:array[@key = ('Authors', 'Editors', 'Collaborators')]/fn:map
                      |Authors/Person
                      |Collaborators/Person
                      |Editors/Person"
                mode="citavi">
    <xsl:param name="person-group-name" as="xs:string?" select="(../@key, parent::*/local-name())[1]"/>
    <xsl:if test="$debug = 'yes'">
      <xsl:comment select="../@key, fn:string[@key = '$id']"/>
    </xsl:if>
    <xsl:choose>
      <xsl:when test="$person-group-name = ('Authors', 'Editors')">
        <xsl:element name="{lower-case(replace($person-group-name, 's$', ''))}">
          <xsl:call-template name="docx2hub:citavi-personname"/>
        </xsl:element>
      </xsl:when>
      <xsl:otherwise>
        <othercredit role="collaborator">
          <xsl:call-template name="docx2hub:citavi-personname"/>
        </othercredit>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="docx2hub:citavi-personname">
    <personname>
      <xsl:apply-templates select="fn:string[@key = 'FirstName'],
                                   fn:string[@key = 'MiddleName'],
                                   fn:string[@key = 'Particle'],
                                   fn:string[@key = 'LastName'],
                                   FirstName,
                                   MiddleName,
                                   Particle,
                                   LastName" 
                           mode="#current"/>
    </personname>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'FirstName']
                      |FirstName" 
                mode="citavi">
    <firstname>
      <xsl:value-of select="."/>
    </firstname>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'LastName']
                      |LastName" 
                mode="citavi">
    <surname>
      <xsl:value-of select="."/>
    </surname>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'MiddleName']
                      |MiddleName"
                mode="citavi">
    <othername role="middle">
      <xsl:value-of select="."/>
    </othername>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'Particle']
                      |Particle" 
                mode="citavi">
    <othername role="particle">
      <xsl:value-of select="."/>
    </othername>
  </xsl:template>
    
  <xsl:template match="fn:string[@key = ('Edition', 'Title')]
                      |Edition
                      |Title"
                mode="citavi">
    <xsl:element name="{lower-case((@key, local-name())[1])}">
      <xsl:value-of select="."/>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = ('Year', 'Date')]
                      |Year
                      |Date" mode="citavi">
    <pubdate role="{lower-case((@key, local-name())[1])}">
      <xsl:value-of select="."/>
    </pubdate>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = ('Number')]
                      |Number" 
                mode="citavi">
    <issuenum>
      <xsl:value-of select="."/>
    </issuenum>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = ('Volume')]
                      |Volume" 
                mode="citavi">
    <volumenum>
      <xsl:value-of select="."/>
    </volumenum>
  </xsl:template>
    
  <xsl:template match="fn:string[@key = ('Doi', 'Isbn', 'Issn')]
                      |Doi
                      |Isbn
                      |Issn" 
                mode="citavi">
    <biblioid class="{lower-case(lower-case((@key, local-name())[1]))}">
      <xsl:value-of select="."/>
    </biblioid>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = ('PageCount', 'Price')]
                      |PageCount
                      |Price" 
                mode="citavi">
    <bibliomisc role="{lower-case(lower-case((@key, local-name())[1]))}">
      <xsl:value-of select="."/>
    </bibliomisc>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'ShortTitle']" 
                mode="citavi">
    <titleabbrev>
      <xsl:value-of select="."/>
    </titleabbrev>
  </xsl:template>
  
  <xsl:template match="fn:map[@key = 'SeriesTitle']
                      |SeriesTitle"
                mode="citavi">    
      <bibliomisc role="seriestitle">
        <xsl:value-of select="(fn:string[@key eq 'Name'], Name)[1]"/>
      </bibliomisc>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'SourceOfBibliographicInformation']
                      |SourceOfBibliographicInformation" 
                mode="citavi">
    <bibliosource>
      <xsl:value-of select="."/>
    </bibliosource>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'Subtitle']
                      |Subtitle" 
                mode="citavi">
    <subtitle>
      <xsl:value-of select="."/>
    </subtitle>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'OnlineAddress']
                      |OnlineAddress" mode="citavi">
    <xsl:attribute name="xlink:href" select="."/>
  </xsl:template>
  
  <xsl:template match="fn:string[@key = 'PageRange']
                      |PageRange" 
                mode="citavi">
    <xsl:variable name="parsed" as="document-node(element(PageRange))" 
                  select="parse-xml('&lt;PageRange>' || . || '&lt;/PageRange>')"/>
    <pagenums>
      <xsl:value-of select="replace($parsed/PageRange/os, '-', '–')"/>
    </pagenums>
  </xsl:template>
  
  <xsl:template match="fn:array[@key = 'Publishers']/fn:map
                      |Publishers" 
                mode="citavi">
    <xsl:apply-templates select="../fn:string[@key = 'PlaceOfPublication'], 
                                 fn:string[@key = 'Name'],
                                 ../PlaceOfPublication,
                                 Publisher/Name" mode="#current"/>
  </xsl:template>
    
  <xsl:template match="fn:array[@key = 'Publishers']/fn:map/fn:string[@key = 'Name']
                      |Publisher/Name" mode="citavi">
    <publishername>
      <xsl:value-of select="."/>
    </publishername>
  </xsl:template>
    
  <xsl:template match="fn:string[@key = 'PlaceOfPublication']
                      |PlaceOfPublication" 
                mode="citavi">
    <address>
      <city>
        <!-- ';' → ' and ' replacement probably necessary for BibTeX --> 
        <xsl:value-of select="."/>
      </city>
    </address>
  </xsl:template>


  <!-- CSL reference manager -->

  <xsl:template match="CSL_JSON" mode="wml-to-dbk tables" 
                use-when="xs:decimal(system-property('xsl:version')) ge 3.0">
    <xsl:param name="csl-refs" as="document-node()?" tunnel="yes"/>
    <xsl:variable name="csl-citation-id" select="." as="xs:string"/>
    <xsl:choose>
      <xsl:when test="$csl-refs">
        <citation linkends="csl-{generate-id()}">
          <xsl:apply-templates mode="wml-to-dbk"/>
        </citation>    
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates mode="wml-to-dbk"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="docx2hub:csl-json-to-xml" use-when="xs:decimal(system-property('xsl:version')) ge 3.0">
    <xsl:variable name="jsons" as="item()*">
      <xsl:try>
        <xsl:for-each select=".//CSL_JSON/@fldArgs">
          <docx2hub:csl-json id="csl-{generate-id(..)}">
            <xsl:sequence select="json-to-xml(.)"/>
          </docx2hub:csl-json>
        </xsl:for-each>
        <xsl:catch/>
      </xsl:try>
    </xsl:variable>
    <xsl:if test="exists($jsons)">
      <xsl:document>
        <docx2hub:csl-jsons>
          <xsl:sequence select="$jsons"/>
        </docx2hub:csl-jsons>
      </xsl:document>
    </xsl:if>
  </xsl:template>

  <xsl:variable name="csl-rendered-by-pos-method" as="xs:string"
    select="'authoryear'"/>

  <xsl:template match="fn:*[@key = 'citationItems']//fn:map[@key = 'itemData']" mode="csl">
    <xsl:variable name="formattedCitation" as="element()?"
      select="ancestor::fn:*[@key = 'citationItems']/parent::*/fn:*[@key = 'properties']/fn:*[@key = 'formattedCitation']"/>
    <xsl:variable name="position-in-CSL-formatted" as="element()">
      <xsl:element name="result">
        <xsl:choose>
          <xsl:when test="$csl-rendered-by-pos-method = 'authoryear'">
            <xsl:variable name="citation-structured" as="node()*">
              <xsl:element name="structured-authoryear-infos">
                <xsl:choose>
                  <xsl:when test="$csl-rendered-by-pos-method = 'authoryear' and
                                  descendant::fn:map[@key = 'issued']/fn:array[@key = 'date-parts']/fn:array/*[1][matches(., '^(19|20)\d\d$')] and
                                  descendant::fn:array[@key = 'author']/fn:map[1]/fn:string[@key = 'family'][normalize-space()]">
                    <xsl:variable name="surname" as="xs:string?"
                      select="normalize-space(descendant::fn:array[@key = 'author']/fn:map[1]/fn:string[@key = 'family'][1])"/>
                    <surname regex-normalized="{replace($surname, '([\{\}\[\]\(\)])', '\\$1')}">
                      <xsl:sequence select="$surname"/>
                    </surname>
                    <xsl:variable name="year" as="xs:string?"
                      select="normalize-space(descendant::fn:map[@key = 'issued']/fn:array[@key = 'date-parts']/fn:array/*[1])"/>
                    <year regex-normalized="{replace($year, '([\{\}\[\]\(\)])', '\\$1')}">
                      <xsl:sequence select="$year"/>
                    </year>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:analyze-string select="replace(normalize-space($formattedCitation), '^\s+|\set\sal\.|\(|\)|,', '')" regex="^[-\p{{L}}]+">
                      <xsl:matching-substring>
                        <surname>
                          <xsl:sequence select="."/>
                        </surname>
                      </xsl:matching-substring>
                      <xsl:non-matching-substring>
                        <xsl:analyze-string select="." regex="\d\d\d\d[a-z]?">
                          <xsl:matching-substring>
                            <year>
                              <xsl:sequence select="."/>
                            </year>
                          </xsl:matching-substring>
                          <xsl:non-matching-substring/>
                        </xsl:analyze-string>
                      </xsl:non-matching-substring>
                    </xsl:analyze-string>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:element>
            </xsl:variable>
            <xsl:variable name="match-candidate" as="element()*"
              select="$root//*:CSL_XML/w:p[
                            $citation-structured/*:year[1] != ''
                        and matches(normalize-space(.), concat('[\s;,.\(]', $citation-structured/*:year[1]/@regex-normalized, '[\s;,.\)]'))
                        and $citation-structured/*:surname != ''
                        and matches(normalize-space(.), concat('^', $citation-structured/*:surname/@regex-normalized))
                      ]"/>
            <xsl:choose>
              <xsl:when test="count($match-candidate) = 1">
                <pos>
                  <xsl:sequence select="$root//*:CSL_XML/w:p[. is $match-candidate]/count(preceding-sibling::w:p) + 1"/>
                </pos>
              </xsl:when>
              <xsl:otherwise/>
            </xsl:choose>
            <xsl:sequence select="$citation-structured"/>
          </xsl:when>
          <!-- xsl:when for other methods -->
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:element>
    </xsl:variable>

    <biblioentry xml:id="{ancestor::docx2hub:csl-json/@id}">
      <xsl:attribute name="docx2hub:rendered-match-searchterm" select="$position-in-CSL-formatted/*:structured-authoryear-infos/*"/>
      <xsl:if test="$position-in-CSL-formatted/*:pos castable as xs:integer">
        <xsl:attribute name="docx2hub:rendered-match-candidate-pos" select="$position-in-CSL-formatted/*:pos"/>
      </xsl:if>
      <xsl:call-template name="csl-reference"/>
      <xsl:if test="contains($debug-dir-uri, 'debug-json-to-xml-bibliography=yes')">
        <docx2hub:debug role="input">
          <xsl:sequence select="."/>
        </docx2hub:debug>
      </xsl:if>
    </biblioentry>
  </xsl:template>

  <xsl:template match="*:CITAVI_XML/w:p | *:CSL_XML/w:p" mode="wml-to-dbk tables">
    <xsl:param name="is-bibliomixed" select="true()" as="xs:boolean" tunnel="yes"/>
    <xsl:choose>
      <xsl:when test="$is-bibliomixed and @role= 'CitaviBibliographyHeading'">
        <title>
          <xsl:apply-templates select="@*, node()" mode="#current"/>
        </title>
      </xsl:when>
      <xsl:when test="$is-bibliomixed">
        <bibliomixed>
          <xsl:apply-templates select="@*, node()" mode="#current"/>
        </bibliomixed>
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="csl-reference">
    <xsl:variable name="reference-type" as="attribute(relation)?">
      <xsl:apply-templates select="fn:string[@key = 'type']" mode="#current"/>
    </xsl:variable>
    <biblioset>
      <xsl:sequence select="$reference-type"/>
      <xsl:apply-templates select="fn:string[@key = 'language']" mode="#current"/>

      <xsl:variable name="authorgroup" as="node()*">
        <xsl:apply-templates select="fn:array[@key = $author-role-values]
          /fn:map" mode="#current"/>
      </xsl:variable>
      <xsl:if test="exists($authorgroup)">
        <authorgroup>
          <xsl:sequence select="$authorgroup"/>
        </authorgroup>
      </xsl:if>

      <xsl:variable name="publisher" as="element(*)*">
        <xsl:apply-templates select="fn:string[@key = ('publisher', 'publisher-place')]" mode="#current"/>
      </xsl:variable>
      <xsl:if test="exists($publisher)">
        <publisher>
          <xsl:sequence select="$publisher"/>
        </publisher>
      </xsl:if>

      <xsl:variable name="conference" as="element(*)*">
        <xsl:apply-templates select="fn:string[@key = ('event', 'event-title', 'event-place', 'event-date')]" mode="#current"/>
      </xsl:variable>
      <xsl:if test="exists($conference)">
        <confgroup>
          <xsl:sequence select="$conference"/>
        </confgroup>
      </xsl:if>

      <xsl:apply-templates select="fn:*[not(@key = (
        (: authorgroup element:)
        $author-role-values,
        
        (: publisher element :)
        'publisher', 'publisher-place', 

        (: conference :)
        'event', 'event-title', 'event-place', 'event-date',

        (: attributes :)
        'type', 'language'))]" mode="#current"/>
    </biblioset>
  </xsl:template>

  <xsl:variable name="author-role-values" as="xs:string*"
    select="('author', 'chair', 'collection-editor', 'compiler', 'composer', 
             'container-author', 'contributor', 'curator', 'director', 'editor', 
             'editorial-director', 'editortranslator', 'executive-producer', 'guest', 
             'host', 'illustrator', 'interviewer', 'narrator', 'organizer', 
             'original-author', 'performer', 'producer', 'recipient', 'reviewed-author', 
             'script-writer', 'series-creator', 'translator')"/>

  <xsl:template match="fn:string[@key = 'type']" mode="csl" as="attribute(relation)?">
    <!-- CSL 1.0.1: article, article-magazine, article-newspaper, article-journal, bill, 
    book, broadcast, chapter, dataset, entry, entry-dictionary, entry-encyclopedia, figure, 
    graphic, interview, legislation, legal_case, manuscript, map, motion_picture, musical_score, 
    pamphlet, paper-conference, patent, post, post-weblog, personal_communication, report, review, review-book, song, speech, thesis, treaty, webpage -->
    <xsl:attribute name="relation" select="."/>
  </xsl:template>

  <xsl:template match="fn:string[@key = ('DOI', 'ISBN', 'ISSN')]" 
                mode="csl">
    <biblioid class="{lower-case(lower-case((@key, local-name())[1]))}">
      <xsl:value-of select="."/>
    </biblioid>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'language']" 
                mode="csl" as="attribute(xml:lang)">
    <xsl:attribute name="xml:lang" select="string(.)"/>
  </xsl:template>

  <xsl:template match="fn:array[@key = $author-role-values]/fn:map" mode="csl">
    <xsl:variable name="el-name" as="xs:string"
      select="(parent::*/@key[. = ('author', 'editor')], 'othercredit')[1]"/>
    <xsl:element name="{$el-name}">
      <xsl:if test="$el-name = 'othercredit'">
        <xsl:attribute name="role" select="parent::*/@key"/>
      </xsl:if>
      <personname>
        <xsl:apply-templates mode="#current"/>
      </personname>
    </xsl:element>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'family']" mode="csl">
    <surname>
      <xsl:apply-templates mode="#current"/>
    </surname>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'given']" mode="csl">
    <firstname>
      <xsl:apply-templates mode="#current"/>
    </firstname>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'suffix']" mode="csl">
    <honorific>
      <xsl:apply-templates mode="#current"/>
    </honorific>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'publisher']" mode="csl">
    <publishername>
      <xsl:apply-templates mode="#current"/>
    </publishername>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'publisher-place']" mode="csl">
    <address>
      <xsl:apply-templates mode="#current"/>
    </address>
  </xsl:template>

  <xsl:template match="fn:string[@key = ('event', 'event-title')]" mode="csl">
    <conftitle>
      <xsl:apply-templates mode="#current"/>
    </conftitle>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'event-place']" mode="csl">
    <address>
      <xsl:apply-templates mode="#current"/>
    </address>
  </xsl:template>

  <!--<xsl:template match="fn:string[@key = 'note']" mode="csl">
    <annotation>
      <xsl:apply-templates mode="#current"/>
    </annotation>
  </xsl:template>-->

  <xsl:template match="fn:*[@key = 'id']" mode="csl">
    <xsl:processing-instruction name="tr_csl_id" select="."/>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'schema']" mode="csl"/>
  <xsl:template match="fn:string[@key = 'formattedCitation']" mode="csl"/>
  <xsl:template match="fn:string[@key = 'manualFormatting']" mode="csl"/>
  <xsl:template match="fn:string[@key = 'plainTextFormattedCitation']" mode="csl"/>
  <xsl:template match="fn:string[@key = 'previouslyFormattedCitation']" mode="csl"/>
  <xsl:template match="fn:string[@key = 'properties']" mode="csl"/>

  <xsl:template match="*[@key = ('comma-suffix', 'dropping-particle', 'parse-names', 'non-dropping-particle', 'static-ordering', 'literal')]" mode="csl">
    <xsl:processing-instruction name="tr_csl_{@key}" select="."/>
  </xsl:template>

  <xsl:template match="fn:array[@key = 'citationItems'] | fn:array[@key = 'citationItems']/fn:map" mode="csl" priority="-1">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="fn:string[not(*)][not(normalize-space())]" mode="csl" priority="1"/>

  <xsl:template match="fn:string[@key = 'title'
                                 or
                                 (@key = 'container-title' and not(../fn:string[@key = 'title']))
                                ]" mode="csl">
    <title>
      <xsl:apply-templates mode="#current"/>
    </title>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'title-short']" mode="csl">
    <titleabbrev>
      <xsl:apply-templates mode="#current"/>
    </titleabbrev>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'source']" mode="csl">
    <bibliosource>
      <xsl:apply-templates mode="#current"/>
    </bibliosource>
  </xsl:template>

  <xsl:template match="fn:*[@key = 'issue']" mode="csl">
    <issuenum>
      <xsl:apply-templates mode="#current"/>
    </issuenum>
  </xsl:template>

  <xsl:template match="fn:*[@key = 'volume']" mode="csl">
    <volumenum>
      <xsl:apply-templates mode="#current"/>
    </volumenum>
  </xsl:template>

  <xsl:template match="fn:*[@key = 'page']" mode="csl">
    <pagenums>
      <xsl:value-of select="replace(., '-', '–')"/>
    </pagenums>
  </xsl:template>

  <xsl:variable name="date-variables" as="xs:string+"
    select="('accessed', 'available-date', 'event-date', 'issued', 'original-date', 'submitted')"/>

  <xsl:template match="fn:*[@key = $date-variables[not(. = 'event-date')]]" mode="csl">
    <releaseinfo role="{@key}">
      <xsl:apply-templates mode="#current"/>
    </releaseinfo>
  </xsl:template>
  <xsl:template match="fn:*[@key = 'event-date']" mode="csl">
    <confdates>
      <xsl:apply-templates mode="#current"/>
    </confdates>
  </xsl:template>

  <xsl:template match="fn:*[@key = $date-variables]//*" mode="csl">
    <phrase role="{(@key, local-name())[1]}">
      <xsl:apply-templates mode="#current"/>
    </phrase>
  </xsl:template>

  <xsl:template match="fn:*[@key = $date-variables]//*[@key = 'date-parts' or self::fn:array]
                                                      [count(.//text()) = 1]" mode="csl" priority="1">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="fn:string[@key = 'abstract']" mode="csl">
    <abstract>
      <para>
        <xsl:apply-templates mode="#current"/>
      </para>
    </abstract>
  </xsl:template>

  <xsl:template mode="csl" priority="-.5" as="element(dbk:bibliomisc)"
    match="fn:*[@key = 'itemData']/fn:string[@key]">
    <bibliomisc role="{@key}">
      <xsl:apply-templates mode="#current"/>
    </bibliomisc>
  </xsl:template>

  <xsl:template match="fn:*[@key]" priority="-3" mode="csl" as="item()*">
    <phrase role="csl_{@key}">
      <xsl:apply-templates mode="#current"/>
    </phrase>
  </xsl:template>

  <xsl:template match="*:biblioset[@xml:lang/normalize-space(.)]/*:language" mode="docx2hub:join-runs"/>

  <xsl:variable name="docx2hub:bibref-id-prefix" select="'bib'" as="xs:string"/>

  <xsl:template match="*:citation/@*[name() = ('linkends', 'linkend')]" mode="docx2hub:join-runs">
    <xsl:attribute name="{name()}" separator=" ">
      <xsl:apply-templates select="key('by-id', tokenize(., '\s+'))/@xml:id" mode="#current"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template match="*:anchor[@role = 'docx2hub:citavi-rendered']/@xml:id" mode="docx2hub:join-runs">
    <xsl:apply-templates select="ancestor::*:biblioentry/@xml:id" mode="#current"/>
  </xsl:template>

  <xsl:key name="docx2hub:by-citavi-citation-linkend" match="*:citation" 
    use="@docx2hub:citavi-rendered-linkend"/>

  <xsl:template match="*[*:bibliography/@role = 'CSL']
                        /*:bibliography[@role = 'CSL-formatted']
                                       [count(*:bibliomixed)
                                      = count(distinct-values(../*:bibliography[@role = 'CSL']/*/@docx2hub:rendered-match-candidate-pos))]
                                       /*:bibliomixed" mode="docx2hub:join-runs">
    <xsl:variable name="rendered-pos" select="count(preceding-sibling::*:bibliomixed) + 1" as="xs:integer"/>
    <xsl:variable name="corresponding-biblioentry" as="element()*"
      select="../../*:bibliography[@role = 'CSL']/*:biblioentry[@docx2hub:rendered-match-candidate-pos/xs:integer(.) = $rendered-pos]"/>
    <biblioentry>
      <xsl:apply-templates select="$corresponding-biblioentry[1]/@*" mode="#current"/>
      <abstract role="rendered">
        <para>
          <xsl:apply-templates select="@*, node()" mode="#current"/>
        </para>
      </abstract>
      <xsl:apply-templates select="$corresponding-biblioentry[1]/node()" mode="#current"/>
    </biblioentry>
  </xsl:template>
  <xsl:template match="*[*:bibliography/@role = 'CSL-formatted']
                        /*:bibliography[@role = 'CSL']
                                        [count(../*:bibliography[@role = 'CSL-formatted']/*:bibliomixed)
                                      = count(distinct-values(*/@docx2hub:rendered-match-candidate-pos))]" mode="docx2hub:join-runs"/>
  <xsl:template match="@docx2hub:rendered-match-searchterm" mode="docx2hub:join-runs"/>
  <xsl:template match="@docx2hub:rendered-match-candidate-pos" mode="docx2hub:join-runs"/>

  <xsl:template match="*:biblioentry[
                         *:abstract[
                           .//*:anchor[@role = 'docx2hub:citavi-rendered']
                                      [exists(key('docx2hub:by-citavi-citation-linkend', current()/@xml:id))]
                         ]
                       ]" mode="docx2hub:join-runs">
    <xsl:variable name="biblioentry-for-current-abstract-order" as="element()?"
      select="../*:biblioentry[
                  @xml:id = key('docx2hub:by-citavi-citation-linkend', current()//*:anchor[@role = 'docx2hub:citavi-rendered']/@xml:id)/@linkends/tokenize(., '\s+')
                ]"/>
    <biblioentry>
      <xsl:apply-templates select="$biblioentry-for-current-abstract-order/@xml:id" mode="#current"/>
      <xsl:apply-templates select="*:abstract" mode="#current"/>
      <xsl:apply-templates select="$biblioentry-for-current-abstract-order/node()[not(self::*:abstract)]" mode="#current"/>
    </biblioentry>
  </xsl:template>
  <xsl:template match="*:citation/@docx2hub:citavi-rendered-linkend" mode="docx2hub:join-runs"/>
  <xsl:template match="*:citation/*:anchor[@xml:id/starts-with(., '_CTVP001')]
                       | *:anchor[@role = 'docx2hub:citavi-rendered']" mode="docx2hub:join-runs"/>

  <xsl:template match="*:citation/*:anchor[@role = 'docx2hub:citavi-rendered'][@xml:id/starts-with(., '_CTVP001')]" mode="docx2hub:join-runs"/>

  <xsl:template match="*:bibliography[@role = 'Citavi-formatted'][exists(*:bibliomixed//*:anchor[starts-with(@xml:id, '_CTVL')])]
                       [
                         every $bibentry in *:bibliomixed
                         satisfies $bibentry//*:anchor[@role = 'docx2hub:citavi-rendered']/@xml:id[. = //*:citation[@linkends/tokenize(., '\s+') = //*:bibliography[@role = 'Citavi']/*:biblioentry/@xml:id]/@docx2hub:citavi-rendered-linkend]
                       ]/*:bibliomixed" mode="docx2hub:join-runs">
    <xsl:variable name="biblioentry-for-current-bibliomixed" as="element()?"
      select="//*:bibliography[@role = 'Citavi']/*:biblioentry[
                  @xml:id = //*:citation[@docx2hub:citavi-rendered-linkend = current()//*:anchor[@role = 'docx2hub:citavi-rendered']/@xml:id]/@linkends/tokenize(., '\s+')
                ]"/>
    <biblioentry>
      <xsl:apply-templates select="$biblioentry-for-current-bibliomixed/@xml:id" mode="#current"/>
      <abstract role="rendered">
        <para>
          <xsl:apply-templates select="@* except @xml:id, node()" mode="#current"/>
        </para>
      </abstract>
      <xsl:apply-templates select="$biblioentry-for-current-bibliomixed/node()" mode="#current"/>
    </biblioentry>
  </xsl:template>
  <xsl:template match="*:bibliography[@role = 'Citavi'][*:biblioentry/@xml:id]
                       [
                         every $bibentry in *:biblioentry
                         satisfies $bibentry//@xml:id = //*:citation[@docx2hub:citavi-rendered-linkend = //*:bibliography[@role = 'Citavi-formatted']/*:bibliomixed//*:anchor[@role = 'docx2hub:citavi-rendered']/@xml:id]/@linkends/tokenize(., '\s+')
                       ]" mode="docx2hub:join-runs"/>

  <xsl:key name="bibentry-by-string" match="*:biblioentry" 
    use="normalize-space(string-join(descendant::text(), ''))"/>

  <xsl:variable name="biblioentry-ids" as="xs:string*"
                select="for $bibentry in //*:biblioentry 
                        return ($bibentry/@xml:id, generate-id($bibentry))[1]"/>

  <xsl:template match="*:biblioentry/@xml:id" mode="docx2hub:join-runs">
    <xsl:variable name="normalized-text" as="xs:string"
      select="normalize-space(string-join(parent::*:biblioentry/descendant::text(), ''))"/>
    <xsl:variable name="current-id" as="attribute(xml:id)" select="."/>
    <xsl:variable name="current-bibentry" as="element()"
      select="if(some $be in parent::*:biblioentry/preceding-sibling::*:biblioentry 
                 satisfies $be[normalize-space(string-join(descendant::text(), '')) = $normalized-text]) 
                then key('bibentry-by-string', $normalized-text)[1] 
                else parent::*:biblioentry"/>
    <xsl:variable name="id-duplicates" as="xs:integer"
                  select="count(parent::*:biblioentry/preceding-sibling::*:biblioentry[@xml:id eq $current-id])" />
    <xsl:attribute name="xml:id" 
      select="concat(
                $docx2hub:bibref-id-prefix, 
                index-of($biblioentry-ids, (., generate-id($current-bibentry))[1])[1],
                concat('_', $id-duplicates + 1)[$id-duplicates gt 0]
              )"/>
  </xsl:template>
  
</xsl:stylesheet>
