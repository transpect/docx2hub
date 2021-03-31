<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:w= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  exclude-result-prefixes="dbk docx2hub xs fn xlink w"
  xmlns="http://docbook.org/ns/docbook"
  version="3.0">

  
  <xsl:template match="CITAVI_JSON" mode="wml-to-dbk tables" 
                use-when="xs:decimal(system-property('xsl:version')) ge 3.0">
    <xsl:param name="citavi-refs" as="document-node()?" tunnel="yes"/>
    <xsl:variable name="citavi-xml" as="element(Placeholder)*"
                  select="for $i in replace(w:bookmarkStart[1]/@w:name, '^_CTVP001', '')
                          return $citavi-refs/docx2hub:citavi-xml/Placeholder[replace(Id, '-', '') eq $i]"/>
    <xsl:choose>
      <xsl:when test="ancestor::w:sdt[1]/w:sdtPr/w:tag/@w:val">
        <biblioref>
          <xsl:apply-templates select="ancestor::w:sdt[1]/w:sdtPr/w:tag/@w:val" mode="wml-to-dbk"/>
          <xsl:comment select="."/>
        </biblioref>    
      </xsl:when>
      <xsl:when test="$citavi-xml">
        <biblioref linkends="{$citavi-xml/Entries/Entry/ReferenceId/concat('_', .)}">
          <xsl:comment select="."/>
        </biblioref>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates mode="wml-to-dbk"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:sdt[.//*:CITAVI_XML]
                      |*:CITAVI_XML" mode="wml-to-dbk tables" priority="2" 
                use-when="xs:decimal(system-property('xsl:version')) ge 3.0">
  </xsl:template>
  

  <xsl:template match="w:sdtPr/w:tag/@w:val[matches(., '^Citavi\.?Placeholder', 'i')]" mode="wml-to-dbk">
    <xsl:param name="citavi-refs" as="document-node()?" tunnel="yes"/>
    <xsl:if test="exists($citavi-refs/docx2hub:citavi-jsons)">
      <xsl:variable name="cited-refs" as="element(fn:map)*" 
        select="key('docx2hub:by-citavi-placeholder', ., $citavi-refs)/fn:array[@key = 'Entries']/fn:map/fn:map[@key = 'Reference']"/>
      <xsl:attribute name="linkends" separator=" " 
        select="for $cid in $cited-refs/fn:string[@key = 'Id'] return '_' || $cid" />
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
      <xsl:try select="for $jd in .//CITAVI_JSON/@fldArgs 
                       return json-to-xml(unparsed-text($jd))">
        <xsl:catch/>
      </xsl:try>
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
  
  <xsl:template match="node() | @*" mode="citavi csl">
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
    <biblioentry xml:id="_{fn:string[@key = 'Id']}">
      <xsl:apply-templates mode="#current" 
        select="fn:map[@key = 'ParentReference'][*/@key]"/>
      <xsl:call-template name="citavi-reference"/>
    </biblioentry>
  </xsl:template>
  
  <xsl:template match="Reference" mode="citavi">
    <biblioentry xml:id="_{Id}">
      <xsl:apply-templates select="ParentReference" mode="#current"/>
      <xsl:call-template name="citavi-reference-xml"/>
    </biblioentry>
  </xsl:template>
  
  <xsl:variable name="other-citavi-ref-parts" as="xs:string+"
                select="'Date',
                        'Doi', 
                        'Edition', 
                        'Isbn', 
                        'Language',
                        'Number',
                        'PageRange',
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
      <xsl:apply-templates select="fn:string[@key = $other-citavi-ref-parts]" mode="#current"/>
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
        <biblioref linkends="csl-{generate-id()}">
          <xsl:apply-templates mode="wml-to-dbk"/>
        </biblioref>    
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

  <xsl:template match="fn:map[@key = 'itemData']" mode="csl">
    <biblioentry xml:id="{ancestor::docx2hub:csl-json/@id}">
      <xsl:call-template name="csl-reference"/>
      <xsl:if test="contains($debug-dir-uri, 'debug-json-to-xml-bibliography=yes')">
        <docx2hub:debug role="input">
          <xsl:sequence select="."/>
        </docx2hub:debug>
      </xsl:if>
    </biblioentry>
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
  
</xsl:stylesheet>
