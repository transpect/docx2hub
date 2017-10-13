<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" 
  xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:tr="http://transpect.io"
  version="1.0" 
  name="mathtype2mml"
  type="docx2hub:mathtype2mml">

  <p:input port="source" primary="true">
    <p:documentation>The result of docx2hub:single-tree (or apply-changemarkup)</p:documentation>
  </p:input>
  <p:input port="zip-manifest">
    <p:documentation>The zip-manifest output port of docx2hub:single-tree</p:documentation>
  </p:input>
  <p:input port="params">
    <p:documentation>The params output port of docx2hub:single-tree</p:documentation>
  </p:input>
  <p:input port="schematron">
    <p:document href="../sch/mathtype2mml.sch.xml"/>
  </p:input>
  <p:input port="custom-font-maps" primary="false" sequence="true">
    <p:documentation xmlns="http://www.w3.org/1999/xhtml">
      <p>A sequence of &lt;symbols&gt; documents, containing mapped characters as found in the regular docs2hub fontmaps.</p>
      <p>Each &lt;symbols&gt; is required to contain the name of its font-family as an attribute @name.</p>
      <p>Example, the value of @char is the unicode character that will be in the mml output:</p>
      <pre>&lt;symbols name="Times New Roman">
  &lt;symbol number="002F" entity="&#x002f;" char="&#x002f;"/>
&lt;/symbols></pre>
      <p>If the base name of the base URI’s file name part is not identical with the font name as encoded in MTEF, 
      you need to give the converter a hint by adding an attribute <code>/symbols/@mathtype-name</code>.</p>
    </p:documentation>
    <p:empty/>
  </p:input>

  <p:output port="result" primary="true">
    <p:documentation>The same basic structure as the primary source of the current step, but with equation OLE objects replaced with MathML</p:documentation>
    <p:pipe port="result" step="convert-mathtype2mml"/>
  </p:output>
  <p:output port="modified-zip-manifest">
    <p:pipe port="zip-manifest" step="convert-mathtype2mml"/>
  </p:output>
  <p:output port="report" sequence="true">
    <p:pipe port="report" step="convert-mathtype2mml"/>
  </p:output>

  <p:serialization port="result" omit-xml-declaration="false"/>
  
  <p:option name="debug" required="false" select="'no'"/>
  <p:option name="debug-dir-uri" required="false" select="'file:/tmp/debug'"/>
  <p:option name="active" required="false" select="'yes'">
    <p:documentation>see corresponding documentation for docx2hub</p:documentation>
  </p:option>
  <p:option name="sources" required="false" select="$mathtype2mml">
    <p:documentation>see documentation for 'active' in docx2hub</p:documentation>
  </p:option>
  <p:option name="mml-space-handling" select="'mspace'">
    <p:documentation>see corresponding documentation for docx2hub</p:documentation>
  </p:option>
  
  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>
  <p:import href="http://transpect.io/calabash-extensions/mathtype-extension/xpl/mathtype2mml-declaration.xpl"/>
  <p:import href="http://transpect.io/xproc-util/store-debug/xpl/store-debug.xpl"/>

  <p:choose name="convert-mathtype2mml">
    <p:when test="not($active eq 'no')">
      <p:output port="result" primary="true"/>
      <p:output port="zip-manifest">
        <p:pipe port="result" step="remove-converted-objects-from-zip-manifest"/>
      </p:output>
      <p:output port="report" sequence="true">
        <p:pipe port="result" step="check"/>
        <p:pipe port="result" step="extract-errors"/>
      </p:output>
      <p:variable name="basename" select="replace(/w:root/@local-href, '^.+/(.+)\.do[ct][mx]$', '$1')"/>
      <p:viewport
        match="/w:root/*[local-name() = ('document', 'footnotes', 'endnotes', 'comments')]//w:object[o:OLEObject[@Type eq 'Embed' and starts-with(@ProgID, 'Equation')]]"
        name="mathtype2mml-viewport">
        <p:variable name="rel-wmf-id" select="w:object/v:shape/v:imagedata/@r:id"
          xmlns:v="urn:schemas-microsoft-com:vml"/>
        <p:variable name="rel-ole-id" select="w:object/o:OLEObject/@r:id"/>
        <p:variable name="rels-elt"
          select="if (contains(base-uri(/*), '/word/document'))
                    then 'w:docRels'
                    else if (contains(base-uri(/*), '/word/footnotes'))
                      then 'w:footnoteRels'
                    else if (contains(base-uri(/*), '/word/endnotes'))
                      then 'w:endnoteRels'
                    else if (contains(base-uri(/*), '/word/comments'))
                      then 'w:commentRels'
                    else ''"/>
        <p:variable name="equation-wmf-href"
          select="if ($rel-wmf-id)
                    then concat(/w:root/@xml:base, 'word/',
                                /w:root/*[name() = $rels-elt]/rel:Relationships/rel:Relationship[@Id eq $rel-wmf-id]/@Target
                               )
                    else 'no-image-found'">
          <p:pipe port="source" step="mathtype2mml"/>
        </p:variable>
        <p:variable name="equation-ole-href"
          select="concat(/w:root/@xml:base, 'word/',
                         /w:root/*[name() = $rels-elt]/rel:Relationships/rel:Relationship[@Id eq $rel-ole-id]/@Target
                        )">
          <p:pipe port="source" step="mathtype2mml"/>
        </p:variable>
        <cx:message>
          <p:with-option name="message"
            select="'wmf:', $rel-wmf-id, ' ole:', $rel-ole-id, ' wmf-href:', $equation-wmf-href"/>
        </cx:message>
        <p:try>
          <p:group>
            <tr:mathtype2mml name="convert-wmf">
              <p:input port="additional-font-maps">
                <p:document href="http://transpect.io/fontmaps/MT_Extra.xml"/>
                <p:document href="http://transpect.io/fontmaps/Symbol.xml"/>
                <p:document href="http://transpect.io/fontmaps/Webdings.xml"/>
                <p:document href="http://transpect.io/fontmaps/Wingdings.xml"/>
                <p:document href="http://transpect.io/fontmaps/Wingdings_2.xml"/>
                <p:document href="http://transpect.io/fontmaps/Wingdings_3.xml"/>
                <p:document href="http://transpect.io/fontmaps/Euclid_Extra.xml"/>
                <p:document href="http://transpect.io/fontmaps/Euclid_Fraktur.xml"/>
                <p:document href="http://transpect.io/fontmaps/Euclid_Math_One.xml"/>
                <p:document href="http://transpect.io/fontmaps/Euclid_Math_Two.xml"/>
                <p:pipe port="custom-font-maps" step="mathtype2mml"/>
              </p:input>
              <p:with-option name="href" select="$equation-wmf-href"/>
              <p:with-option name="debug" select="$debug"/>
              <p:with-option name="debug-dir-uri" select="concat($debug-dir-uri, '/docx2hub/', $basename, '/')"/>
            </tr:mathtype2mml>
            <p:choose>
              <p:when test="/c:errors or matches($active, 'ole|yes')">
                <tr:mathtype2mml name="convert-ole">
                  <p:input port="additional-font-maps">
                    <p:document href="http://transpect.io/fontmaps/MT_Extra.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Symbol.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Webdings.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Wingdings.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Wingdings_2.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Wingdings_3.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Euclid_Extra.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Euclid_Fraktur.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Euclid_Math_One.xml"/>
                    <p:document href="http://transpect.io/fontmaps/Euclid_Math_Two.xml"/>
                    <p:pipe port="custom-font-maps" step="mathtype2mml"/>
                  </p:input>
                  <p:with-option name="href" select="$equation-ole-href"/>
                  <p:with-option name="debug" select="$debug"/>
                  <p:with-option name="debug-dir-uri" select="concat($debug-dir-uri, '/docx2hub/', $basename, '/')"/>
                </tr:mathtype2mml>
                <p:choose>
                  <p:when test="/c:errors">
                    <!-- wmf error, ole error -->
                    <p:set-attributes match="/*" name="set-atts">
                      <p:input port="source" select="/c:errors/c:error">
                        <p:pipe port="result" step="convert-ole"/>
                      </p:input>
                      <p:input port="attributes">
                        <p:pipe port="result" step="convert-ole"/>
                      </p:input>
                    </p:set-attributes>
                    <p:add-attribute name="add-srcpath" attribute-name="srcpath" match="/*">
                      <p:with-option name="attribute-value" select="/*/@srcpath">
                        <p:pipe port="current" step="mathtype2mml-viewport"/>
                      </p:with-option>
                    </p:add-attribute>
                    <p:add-attribute name="add-role" attribute-name="role" attribute-value="error" match="/*"/>
                    <p:sink/>
                    <p:identity>
                      <p:input port="source">
                        <p:pipe port="current" step="mathtype2mml-viewport"/>
                        <p:pipe port="result" step="add-role"/>
                      </p:input>
                    </p:identity>
                  </p:when>
                  <p:when test="matches($active, 'wmf')">
                    <!-- wmf ok, ole ok -->
                    <p:compare name="compare-mml" fail-if-not-equal="false">
                      <p:input port="source">
                        <p:pipe port="result" step="convert-wmf"/>
                      </p:input>
                      <p:input port="alternate">
                        <p:pipe port="result" step="convert-ole"/>
                      </p:input>
                    </p:compare>
                    <p:identity>
                      <p:input port="source">
                        <p:pipe port="result" step="compare-mml"/>
                      </p:input>
                    </p:identity>
                    <p:choose>
                      <p:when test="c:result = true()">
                        <!-- wmf equals ole, only output MathML once -->
                        <p:identity>
                          <p:input port="source">
                            <p:pipe port="result" step="convert-wmf"/>
                          </p:input>
                        </p:identity>
                      </p:when>
                      <p:otherwise>
                        <!-- wmf differs from ole, output both and a pi -->
                        <p:choose>
                          <p:when test="matches($active, '^ole')">
                            <p:wrap-sequence wrapper="wrap-mml">
                              <p:input port="source">
                                <p:pipe port="result" step="convert-ole"/>
                                <p:pipe port="result" step="convert-wmf"/>
                              </p:input>
                            </p:wrap-sequence>
                          </p:when>
                          <p:otherwise>
                            <p:wrap-sequence wrapper="wrap-mml">
                              <p:input port="source">
                                <p:pipe port="result" step="convert-wmf"/>
                                <p:pipe port="result" step="convert-ole"/>
                              </p:input>
                            </p:wrap-sequence>
                          </p:otherwise>
                        </p:choose>
                        <p:identity name="merge"/>
                        <p:insert match="*:math[1]" position="after">
                          <p:input port="source">
                            <p:pipe port="result" step="merge"/>
                          </p:input>
                          <p:input port="insertion">
                            <p:inline><wrap-mml><?tr M2M_201 MathML equation (sources: wmf, ole) differ?></wrap-mml></p:inline>
                          </p:input>
                        </p:insert>
                      </p:otherwise>
                    </p:choose>
                  </p:when>
                  <p:otherwise>
                    <!-- wmf ok, ole ok, no diff -->
                    <p:identity>
                      <p:input port="source">
                        <p:pipe port="result" step="convert-ole"/>
                      </p:input>
                    </p:identity>
                  </p:otherwise>
                </p:choose>
              </p:when>
              <p:otherwise>
                <!-- wmf ok, no ole -->
                <p:identity>
                  <p:input port="source">
                    <p:pipe port="result" step="convert-wmf"/>
                  </p:input>
                </p:identity>
              </p:otherwise>
            </p:choose>
            <p:identity name="chosen-mml"/>
            <p:insert match="o:OLEObject[@Type eq 'Embed' and starts-with(@ProgID, 'Equation')]" position="after">
              <p:input port="source">
                <p:pipe port="current" step="mathtype2mml-viewport"/>
              </p:input>
              <p:input port="insertion">
                <p:pipe port="result" step="chosen-mml"/>
              </p:input>
            </p:insert>
            <p:delete match="o:OLEObject[@Type eq 'Embed' and starts-with(@ProgID, 'Equation')]"/>
          </p:group>
          <p:catch>
            <cx:message>
              <p:with-option name="message" select="'catch :(', node()"/>
            </cx:message>
            <p:identity/>
          </p:catch>
        </p:try>
      </p:viewport>
      <p:unwrap match="wrap-mml"></p:unwrap>
      <p:group name="remove-unused-rels">
        <p:output port="result"/>
        <p:choose>
          <p:when
            test="*[name() = ('w:docRels', 'w:footnoteRels', 'w:endnoteRels', 'w:commentRels')]/rel:Relationships/rel:Relationship
                  [@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject']">
            <p:xslt>
              <p:input port="stylesheet">
                <p:inline>
                  <xsl:stylesheet version="2.0">
                    <xsl:template
                      match="*[name() = ('w:docRels', 'w:footnoteRels', 'w:endnoteRels', 'w:commentRels')]/rel:Relationships/rel:Relationship
                             [@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject']">
                      <xsl:variable name="objects"
                        select="if(ancestor::w:docRels) 
                                  then /w:root/w:document//o:OLEObject
                                else if(ancestor::w:footnoteRels)
                                  then /w:root/w:footnotes//o:OLEObject
                                else if(ancestor::w:endnoteRels)
                                  then /w:root/w:endnotes//o:OLEObject
                                else if(ancestor::w:commentRels)
                                  then /w:root/w:comments//o:OLEObject
                                else ()"
                        as="element(o:OLEObject)*"/>
                      <xsl:copy>
                        <xsl:apply-templates select="@*"/>
                        <xsl:if test="not(@Id = $objects/@r:id)">
                          <xsl:attribute name="remove" select="'yes'"/>
                        </xsl:if>
                        <xsl:apply-templates/>
                      </xsl:copy>
                    </xsl:template>
                    <xsl:template match="node() | @*">
                      <xsl:copy>
                        <xsl:apply-templates select="@*, node()"/>
                      </xsl:copy>
                    </xsl:template>
                  </xsl:stylesheet>
                </p:inline>
              </p:input>
              <p:input port="parameters">
                <p:empty/>
              </p:input>
            </p:xslt>
          </p:when>
          <p:otherwise>
            <p:identity/>
          </p:otherwise>
        </p:choose>
      </p:group>

      <tr:store-debug name="store-viewport">
        <p:with-option name="pipeline-step" select="concat('docx2hub/', $basename, '/02b-mathtype-converted')"/>
        <p:with-option name="active" select="$debug"/>
        <p:with-option name="base-uri" select="$debug-dir-uri"/>
      </tr:store-debug>
      
      <p:validate-with-schematron assert-valid="false" name="val-sch">
        <p:input port="schema">
          <p:pipe port="schematron" step="mathtype2mml"/>
        </p:input>
        <p:input port="parameters">
          <p:empty/>
        </p:input>
        <p:with-param name="allow-foreign" select="'true'"/>
      </p:validate-with-schematron>

      <p:sink/>
    
      <p:add-attribute name="check0" match="/*" 
        attribute-name="tr:rule-family" attribute-value="docx2hub_mathtype2mml">
        <p:input port="source">
          <p:pipe port="report" step="val-sch"/>
        </p:input>
      </p:add-attribute>
      
      <p:insert name="check" match="/*" position="first-child">
        <p:input port="insertion" select="/*/*:title">
          <p:pipe port="schematron" step="mathtype2mml"/>
        </p:input>
      </p:insert>
      
      <p:sink/>
      
      <p:identity>
        <p:input port="source">
          <p:pipe port="result" step="store-viewport"/>
        </p:input>
      </p:identity>
      
      <p:choose name="extract-errors">
        <p:when test="exists(//c:error)">
          <p:output port="result" primary="true" sequence="true"/>
          <p:wrap-sequence wrapper="c:errors">
            <p:input port="source" select="//c:error"></p:input>
          </p:wrap-sequence>
          <p:add-attribute match="/*" attribute-name="tr:rule-family" attribute-value="docx2hub_mathtype2mml"/>
        </p:when>
        <p:otherwise>
          <p:output port="result" primary="true" sequence="true"/>
          <p:identity>
            <p:input port="source">
              <p:inline>
                <c:ok tr:rule-family="docx2hub_mathtype2mml"/>
              </p:inline>
            </p:input>
          </p:identity>
        </p:otherwise>
      </p:choose>
      
      <p:sink/>
      
      <p:xslt name="remove-converted-objects-from-zip-manifest">
        <p:input port="source">
          <p:pipe port="zip-manifest" step="mathtype2mml"/>
          <p:pipe port="result" step="remove-unused-rels"/>
        </p:input>
        <p:input port="parameters"><p:empty/></p:input>
        <p:input port="stylesheet">
          <p:inline>
            <xsl:stylesheet version="2.0">
              <xsl:template match="node() | @*">
                <xsl:copy>
                  <xsl:apply-templates select="@*, node()"/>
                </xsl:copy>
              </xsl:template>
              <xsl:variable name="removed-rels" as="element(rel:Relationship)*"
                select="collection()[2]//rel:Relationship[@remove = 'yes']"/>
              <xsl:template match="c:entry">
                <xsl:choose>
                  <xsl:when test="some $t in $removed-rels/@Target satisfies (ends-with(@name, $t))"/>
                  <xsl:otherwise>
                    <xsl:next-match/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:template>
            </xsl:stylesheet>
          </p:inline>
        </p:input>
      </p:xslt>
      
      <tr:store-debug name="store-zip-manifest">
        <p:with-option name="pipeline-step" select="concat('docx2hub/', $basename, '/02c-modified-zip-manifest')"/>
        <p:with-option name="active" select="$debug"/>
        <p:with-option name="base-uri" select="$debug-dir-uri"/>
      </tr:store-debug>
      
      <p:sink/>

      <p:delete match="rel:Relationship[@remove = 'yes'] | c:error" name="remove-rels">
        <p:input port="source">
          <p:pipe port="result" step="store-viewport"/>  
        </p:input>
      </p:delete>

    </p:when>
    <p:otherwise>
      <p:output port="result" primary="true"/>
      <p:output port="report" sequence="true">
        <p:inline>
          <c:ok tr:rule-family="docx2hub_mathtype2mml"/>
        </p:inline>
      </p:output>
      <p:output port="zip-manifest">
        <p:pipe port="zip-manifest" step="mathtype2mml"/>
      </p:output>
      <p:identity/>
    </p:otherwise>
  </p:choose>
  
  <p:sink/>

</p:declare-step>
