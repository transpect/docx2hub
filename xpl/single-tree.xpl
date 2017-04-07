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
  name="docx-single-tree"
  type="docx2hub:single-tree">

  <p:input port="source" primary="true">
    <p:documentation>This is to prevent any other default readable port to be connected with the xslt port.</p:documentation>
    <p:empty/>
  </p:input>
  <p:input port="xslt" primary="false">
    <p:document href="../xsl/main.xsl"/>
  </p:input>
  <p:input port="insert-xpath-schematron">
    <p:document href="../sch/insert-xpath.sch.xml"/>
    <p:documentation>Schematron that will validate the entire Word container document (before mode apply-changemarkup).
    The tr:rule-family for this Schematron is 'docx2hub_single-tree'. We have kept the port and Schematron file names
    (insert-xpath) for compatibility reasons.</p:documentation>
  </p:input>
  <p:input port="single-tree-schematron">
    <p:document href="../sch/single-tree.sch.xml"/>
    <p:documentation>Schematron that will validate the entire Word container document (after mode apply-changemarkup).
    The tr:rule-family for this Schematron is 'docx2hub_single-tree_changes-accepted'</p:documentation>
  </p:input>
  
  <p:output port="result" primary="true">
    <p:pipe port="result" step="add-xml-base-attr"/>
  </p:output>
  <p:output port="params">
    <p:pipe port="result" step="params"/>
  </p:output>
  <p:output port="zip-manifest">
    <p:pipe port="result" step="zip-manifest"/>
  </p:output>
  <p:output port="report" sequence="true">
    <p:pipe port="result" step="check-insert-xpath"/>
    <p:pipe port="result" step="check-single-tree"/>
    <p:pipe port="report" step="insert-xpath"/>
    <p:pipe port="report" step="apply-changemarkup"/>
  </p:output>

  <p:serialization port="result" omit-xml-declaration="false"/>
  
  <p:option name="docx" required="true">
    <p:documentation>OS name (preferably with full path, may not resolve if only a relative path is given), file:, http:, or
      https: URL. The file will be fetched first if it is an HTTP URL.</p:documentation>
  </p:option>
  <p:option name="debug" required="false" select="'no'"/>
  <p:option name="debug-dir-uri" required="false" select="'file:/tmp/debug'"/>
  <p:option name="srcpaths" required="false" select="'no'"/>
  <p:option name="unwrap-tooltip-links" required="false" select="'no'"/>
  <p:option name="mml-space-handling" select="'mspace'">
    <p:documentation>see corresponding documentation for docx2hub</p:documentation>
  </p:option>
  <p:option name="hub-version" required="false" select="'1.2'"/>
  <p:option name="fail-on-error" select="'no'"/>
  <p:option name="field-vars" required="false" select="'no'"/>
  <p:option name="extract-dir" required="false" select="''">
    <p:documentation>Directory (OS path, not file: URL) to which the file will be unzipped. If option is empty string, will be
      '.tmp' appended to OS file path.</p:documentation>
  </p:option>
  <p:option name="no-srcpaths-for-text-runs-threshold" select="'40000'">
    <p:documentation>In order to speed up conversion for long documents, if more w:r elements are found, they won’t receive 
      a srcpath of their own. In principle, srcpath generation may be sped up by computing them more efficiently,
      building on a tunnelled parameter that contains the parent element’s already-computed srcpath.</p:documentation>
  </p:option>
  <p:option name="apply-changemarkup" required="false" select="'yes'">
    <p:documentation xmlns="http://www.w3.org/1999/xhtml">
      <p>Apply all change markup on the compound word document.</p>
    </p:documentation>
  </p:option>
  
  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>
  
  <p:import href="http://transpect.io/calabash-extensions/unzip-extension/unzip-declaration.xpl"/>
  <p:import href="http://transpect.io/calabash-extensions/mathtype-extension/xpl/mathtype2mml-declaration.xpl"/>
  <p:import href="http://transpect.io/xproc-util/store-debug/xpl/store-debug.xpl"/>
  <p:import href="http://transpect.io/xproc-util/file-uri/xpl/file-uri.xpl"/>
  <p:import href="http://transpect.io/xproc-util/xslt-mode/xpl/xslt-mode.xpl"/>

  <p:variable name="basename" select="replace($docx, '^(.+?)([^/\\]+)\.do[ct][mx]$', '$2')"/>

  <tr:file-uri name="locate-docx">
    <p:with-option name="filename" select="$docx"/>
  </tr:file-uri>
  
  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat('docx2hub/', $basename, '/00-file-uri')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>

  <!-- unzip or error message -->
  <tr:unzip name="unzip">
    <p:with-option name="zip" select="/c:result/@os-path">
      <p:pipe step="locate-docx" port="result"/>
    </p:with-option>
    <p:with-option name="dest-dir"
      select="if ($extract-dir = '') 
              then concat(/c:result/@os-path, '.tmp')
              else $extract-dir">
      <p:pipe step="locate-docx" port="result"/>
    </p:with-option>
    <p:with-option name="overwrite" select="'yes'"/>
  </tr:unzip>

  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat('docx2hub/', $basename, '/00-unzip')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>
  
  <p:choose>
    <p:when test="name(/*) eq 'c:error'">
      <cx:message>
        <p:with-option name="message" select="'docx2hub error on unzipping.&#xa;', //text(), '&#xa;'"/>
      </cx:message>
    </p:when>
    <p:otherwise>
      <p:identity/>
    </p:otherwise>
  </p:choose>

  <p:load name="document">
    <p:with-option name="href" 
      select="concat(
                /c:files/@xml:base,
                (/c:files/c:file/@name[matches(., '^word/document\d?.xml$')])[1]
              )"/>
  </p:load>

  <p:sink/>
  
  <p:xslt name="zip-manifest">
    <p:input port="source">
      <p:pipe port="result" step="unzip"/>
    </p:input>
    <p:input port="stylesheet">
      <p:inline>
        <xsl:stylesheet version="2.0">
          <xsl:template match="c:files">
            <c:zip-manifest>
              <xsl:copy-of select="@xml:base"/>
              <xsl:apply-templates/>
            </c:zip-manifest>
          </xsl:template>
          <xsl:variable name="base-uri" select="/*/@xml:base" as="xs:string"/>
          <xsl:template match="c:file">
            <c:entry name="{replace(replace(@name, '%5B', '['), '%5D', ']')}"
              href="{concat($base-uri, replace(replace(@name, '\[', '%5B'), '\]', '%5D'))}" compression-method="deflate"
              compression-level="default"/>
          </xsl:template>
        </xsl:stylesheet>
      </p:inline>
    </p:input>
    <p:input port="parameters">
      <p:empty/>
    </p:input>
  </p:xslt>
  
  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat('docx2hub/', $basename, '/00-zip-manifest')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>
  
  <p:sink/>

  <p:add-attribute attribute-name="value" match="/c:param" name="error-msg-file-path">
    <p:with-option name="attribute-value" select="replace(static-base-uri(), '/[^/]+$', '')"/>
    <p:input port="source">
      <p:inline><c:param name="error-msg-file-path"/></p:inline>
    </p:input>
  </p:add-attribute>
  
  <p:add-attribute attribute-name="value" match="/c:param" name="local-href">
    <p:with-option name="attribute-value" select="/c:result/@local-href">
      <p:pipe port="result" step="locate-docx"/>
    </p:with-option>
    <p:input port="source">
      <p:inline><c:param name="local-href"/></p:inline>
    </p:input>
  </p:add-attribute>
  
  <p:add-attribute attribute-name="value" match="/c:param" name="srcpaths">
    <p:with-option name="attribute-value" select="$srcpaths"/>
    <p:input port="source">
      <p:inline><c:param name="srcpaths"/></p:inline>
    </p:input>
  </p:add-attribute>

  <p:add-attribute attribute-name="value" match="/c:param" name="extract-dir-uri">
    <p:with-option name="attribute-value" select="/c:files/@xml:base">
      <p:pipe port="result" step="unzip"/>
    </p:with-option>
    <p:input port="source">
      <p:inline><c:param name="extract-dir-uri"/></p:inline>
    </p:input>
  </p:add-attribute>

  <p:in-scope-names name="vars"/>

  <p:insert position="last-child" name="params">
    <p:input port="source">
      <p:pipe port="result" step="vars"/>
    </p:input>
    <p:input port="insertion">
      <p:pipe port="result" step="error-msg-file-path"/>
      <p:pipe port="result" step="extract-dir-uri"/>
      <p:pipe port="result" step="local-href"/>
      <p:pipe port="result" step="srcpaths"/>
    </p:input>
  </p:insert>

  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat('docx2hub/', $basename, '/00-params')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>
  
  <p:sink/>
  
  <tr:xslt-mode msg="yes" mode="insert-xpath" name="insert-xpath">
    <p:input port="source">
      <p:pipe step="document" port="result"/>
    </p:input>
    <p:input port="parameters">
      <p:pipe step="params" port="result"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx-single-tree" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/01')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="srcpaths" select="$srcpaths"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="srcpaths-on-runs" 
      select="if (count(//w:r) &gt; $no-srcpaths-for-text-runs-threshold)
              then 'no' else 'yes'">
      <p:pipe step="document" port="result"/>
    </p:with-param>
  </tr:xslt-mode>

      
  <p:viewport match="/w:root/w:document/w:body/*[local-name() = ('p', 'tbl')]//w:object/o:OLEObject[@Type eq 'Embed' and starts-with(@ProgID, 'Equation')]" name="mathtype2mml-viewport">
    <p:variable name="rel-id" select="o:OLEObject/@r:id"/>
    <p:variable name="equation-href" select="concat(/w:root/@xml:base,
                                                    'word/',
                                                    /w:root/w:docRels/rel:Relationships/rel:Relationship[@Id eq $rel-id]/@Target
                                          )">
      <p:pipe port="result" step="insert-xpath"/>
    </p:variable>
    
    <tr:mathtype2mml name="mathtype2mml">
      <p:with-option name="href" select="$equation-href"/>
      <p:with-option name="debug" select="$debug"/>
      <p:with-option name="debug-dir-uri" select="concat($debug-dir-uri, '/docx2hub/', $basename, '/')"/>
    </tr:mathtype2mml>
    
  </p:viewport>
      
  <tr:store-debug>
    <p:with-option name="pipeline-step" select="concat('docx2hub/', $basename, '/02-mathtype-converted')"/>
    <p:with-option name="active" select="$debug"/>
    <p:with-option name="base-uri" select="$debug-dir-uri"/>
  </tr:store-debug>
  
  <p:choose name="apply-changemarkup">
    <p:when test="exists(//w:del | //w:moveFrom | //w:ins) and $apply-changemarkup = 'yes'">
      <p:output port="result" primary="true"/>
      <p:output port="report" sequence="true">
        <p:pipe port="report" step="apply-changemarkup-xslt"/>
      </p:output>
      
      <tr:xslt-mode msg="yes" mode="docx2hub:apply-changemarkup" name="apply-changemarkup-xslt">
        <p:input port="parameters">
          <p:pipe step="params" port="result"/>
        </p:input>
        <p:input port="stylesheet">
          <p:pipe step="docx-single-tree" port="xslt"/>
        </p:input>
        <p:input port="models">
          <p:empty/>
        </p:input>
        <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/02')"/>
        <p:with-option name="debug" select="$debug"/>
        <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
        <p:with-option name="fail-on-error" select="$fail-on-error"/>
        <p:with-param name="fail-on-error" select="$fail-on-error"/>
      </tr:xslt-mode>
      
    </p:when>
    <p:otherwise>
      <p:output port="result" primary="true"/>
      <p:output port="report" sequence="true">
        <p:empty/>
      </p:output>
      <p:identity/>
    </p:otherwise>
  </p:choose>

  <p:add-attribute attribute-name="xml:base" match="/*" name="add-xml-base-attr">
    <p:with-option name="attribute-value" select="/c:files/@xml:base">
      <p:pipe port="result" step="unzip"/>
    </p:with-option>
  </p:add-attribute>
  
  <p:validate-with-schematron assert-valid="false" name="single-tree0">
    <p:input port="schema">
      <p:pipe port="single-tree-schematron" step="docx-single-tree"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
    <p:with-param name="allow-foreign" select="'true'"/>
  </p:validate-with-schematron>

  <p:sink/>

  <p:add-attribute name="check-single-tree" match="/*" 
    attribute-name="tr:rule-family" attribute-value="docx2hub_single-tree_changes-accepted">
    <p:input port="source">
      <p:pipe port="report" step="single-tree0"/>
    </p:input>
  </p:add-attribute>
  
  <p:sink/>
  
  <p:validate-with-schematron assert-valid="false" name="insert-xpath0">
    <p:input port="source">
      <p:pipe port="result" step="insert-xpath"/>
    </p:input>
    <p:input port="schema">
      <p:pipe port="insert-xpath-schematron" step="docx-single-tree"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
    <p:with-param name="allow-foreign" select="'true'"/>
  </p:validate-with-schematron>
  
  <p:sink/>
  
  <p:add-attribute match="/*" 
    attribute-name="tr:step-name" attribute-value="docx2hub">
    <p:input port="source">
      <p:pipe port="report" step="insert-xpath0"/>
    </p:input>
  </p:add-attribute>

  <p:add-attribute name="check-insert-xpath" match="/*" 
    attribute-name="tr:rule-family" attribute-value="docx2hub_single-tree">
    <p:documentation>Please note that tr:rule-family="docx2hub_single-tree" is checked by insert-xpath-schematron</p:documentation>
  </p:add-attribute>
  
  <p:sink/>

</p:declare-step>
