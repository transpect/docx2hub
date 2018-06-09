<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" 
  xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:tr="http://transpect.io"
  version="1.0" 
  name="single-tree"
  type="docx2hub:single-tree">

  <p:input port="source" primary="true">
    <p:documentation>This is to prevent any other default readable port to be connected with the xslt port.</p:documentation>
    <p:empty/>
  </p:input>
  <p:input port="schematron" primary="false">
    <p:document href="../sch/single-tree.sch.xml"/>
  </p:input>
  <p:input port="xslt" primary="false">
    <p:document href="../xsl/main.xsl"/>
  </p:input>

  <p:output port="result" primary="true">
    <p:pipe port="result" step="group"/>
  </p:output>
  <p:output port="params">
    <p:pipe port="params" step="group"/>
  </p:output>
  <p:output port="zip-manifest">
    <p:pipe port="zip-manifest" step="group"/>
  </p:output>
  <p:output port="report" sequence="true">
    <p:pipe port="report" step="group"/>
  </p:output>
  <p:output port="schema" sequence="true">
    <p:pipe port="schematron-atts" step="group"/>
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
  <p:option name="use-filename-from-http-response" required="false" select="'no'">
    <p:documentation>Use filename that is passed on from http request response instead of 
    possible filename read from URL in tr:file-uri (for example when using Gdocs URLs:
    https://docs.google.com/document/d/1Z5eYyjLoRhB24HYZ-d-wQKAFD3QDWZUsQH4cKHs2eiM/export?format=docx)</p:documentation>
  </p:option>

  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>

  <p:import href="http://transpect.io/calabash-extensions/unzip-extension/unzip-declaration.xpl"/>
  <p:import href="http://transpect.io/xproc-util/store-debug/xpl/store-debug.xpl"/>
  <p:import href="http://transpect.io/xproc-util/file-uri/xpl/file-uri.xpl"/>
  <p:import href="http://transpect.io/xproc-util/xslt-mode/xpl/xslt-mode.xpl"/>

  <tr:file-uri name="locate-docx">
    <p:with-option name="filename" select="$docx"/>
    <p:with-option name="use-filename-from-http-response" select="$use-filename-from-http-response"/>
  </tr:file-uri>

  <p:group name="group">
    <p:output port="result" primary="true">
      <p:pipe port="result" step="insert-archive-uri"/>
    </p:output>
    <p:output port="params">
      <p:pipe port="result" step="params"/>
    </p:output>
    <p:output port="zip-manifest">
      <p:pipe port="result" step="zip-manifest"/>
    </p:output>
    <p:output port="report" sequence="true">
      <p:pipe port="result" step="check"/>
    </p:output>
    <p:output port="schematron-atts" sequence="true">
      <p:pipe port="result" step="schematron-atts"/>
    </p:output>
    
    <p:variable name="basename"
      select="replace(/c:result/@local-href, '^(.+?)([^/\\]+)\.do[ct][mx]$', '$2')">
      <p:pipe port="result" step="locate-docx"/>
    </p:variable>

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

    <p:sink name="sink1"/>

    <p:xslt name="zip-manifest" cx:depends-on="document">
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

    <p:sink name="sink2"/>

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

    <p:sink name="sink3"/>

    <tr:xslt-mode msg="yes" mode="insert-xpath" name="insert-xpath">
      <p:input port="source">
        <p:pipe step="document" port="result"/>
      </p:input>
      <p:input port="parameters">
        <p:pipe step="params" port="result"/>
      </p:input>
      <p:input port="stylesheet">
        <p:pipe step="single-tree" port="xslt"/>
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

    <p:add-attribute attribute-name="xml:base" match="/*" name="add-xml-base-attr">
      <p:with-option name="attribute-value" select="/c:files/@xml:base">
        <p:pipe port="result" step="unzip"/>
      </p:with-option>
    </p:add-attribute>
    
    <p:add-attribute match="/*" name="insert-archive-uri-local" attribute-name="archive-uri-local">
      <p:with-option name="attribute-value" select="/*/@local-href">
        <p:pipe port="result" step="locate-docx"/>
      </p:with-option>
    </p:add-attribute>
    
    <p:add-attribute match="/*" name="insert-archive-uri" attribute-name="archive-uri">
      <p:with-option name="attribute-value" select="/*/@href">
        <p:pipe port="result" step="locate-docx"/>
      </p:with-option>
    </p:add-attribute>

    <p:validate-with-schematron assert-valid="false" name="val-sch">
      <p:input port="schema">
        <p:pipe port="schematron" step="single-tree"/>
      </p:input>
      <p:input port="parameters"><p:empty/></p:input>
      <p:with-param name="allow-foreign" select="'true'"/>
    </p:validate-with-schematron>

    <p:sink name="sink4"/>

    <p:add-attribute name="check0" match="/*" attribute-name="tr:rule-family" attribute-value="docx2hub_single-tree">
      <p:input port="source">
        <p:pipe port="report" step="val-sch"/>
      </p:input>
    </p:add-attribute>

    <p:insert name="check" match="/*" position="first-child">
      <p:input port="insertion" select="/*/*:title">
        <p:pipe port="schematron" step="single-tree"/>
      </p:input>
    </p:insert>
  
  <p:sink name="sink5"/>

  <p:add-attribute match="/*" 
   attribute-name="tr:step-name" 
    attribute-value="docx2hub">
    <p:input port="source">
      <p:pipe port="schematron" step="single-tree"/>
    </p:input>
  </p:add-attribute>
  
  <p:add-attribute 
    match="/*" 
    attribute-name="tr:rule-family" 
    attribute-value="docx2hub">
  </p:add-attribute>
  
  <p:insert name="schematron-atts" match="/*" position="first-child">
    <p:input port="insertion" select="/*/*:title">
      <p:pipe port="schematron" step="single-tree"/>
    </p:input>
  </p:insert>
  
  <p:sink name="sink6"/>

  </p:group>

</p:declare-step>
