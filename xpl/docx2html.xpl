<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" 
  xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions" 
  xmlns:tr="http://transpect.io"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:hub2htm="http://transpect.io/hub2htm" 
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
  version="1.0" 
  name="docx2html" 
  type="docx2hub:docx2html">

  <p:documentation xmlns="http://www.w3.org/1999/xhtml">
    <p>Converts docx to XHTML</p>
  </p:documentation>

  <p:input port="docx2hub-xslt">
    <p:document href="../xsl/main.xsl"/>
  </p:input>
  <p:input port="source" primary="true">
    <p:documentation>This is to prevent a default readable port connecting to this step’s xslt port.</p:documentation>
    <p:empty/>
  </p:input>
  <p:output port="result" primary="true"/>
  <p:serialization port="result" method="xhtml" omit-xml-declaration="false"/>
  <p:output port="hub">
    <p:pipe port="result" step="docx2hub"/>
  </p:output>
  <p:serialization port="hub" omit-xml-declaration="false"/>
  
  <p:option name="docx" required="true">
    <p:documentation>OS name (preferably with full path, may not resolve if only a relative path is given), file:, http:, or
      https: URL. The file will be fetched first if it is an HTTP URL.</p:documentation>
  </p:option>
  <p:option name="debug" required="false" select="'no'"/>
  <p:option name="debug-dir-uri" required="false" select="'debug'"/>
  <p:option name="srcpaths" required="false" select="'no'"/>
  <p:option name="unwrap-tooltip-links" required="false" select="'no'"/>
  <p:option name="hub-version" required="false" select="'1.2'"/>
  <p:option name="fail-on-error" required="false" select="'no'"/>
  <p:option name="extract-dir" required="false" select="''">
    <p:documentation>Directory (OS path, not file: URL) to which the file will be unzipped. If option is empty string, will be
      '.tmp' appended to OS file path.</p:documentation>
  </p:option>
  <p:option name="create-svg" required="false" select="'no'">
    <p:documentation>Whether Office Open Drawing ML should be mapped to SVG</p:documentation>
  </p:option>
  
  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>
  <p:import href="http://transpect.io/hub2html/xpl/hub2html.xpl"/>
  
  <p:import href="docx2hub.xpl"/>
  
  <docx2hub:convert name="docx2hub" charmap-policy="msoffice">
    <p:with-option name="docx" select="$docx"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-option name="unwrap-tooltip-links" select="$unwrap-tooltip-links"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-option name="srcpaths" select="$srcpaths"/>
    <p:with-option name="extract-dir" select="$extract-dir"/>
    <p:with-option name="create-svg" select="$create-svg"/>
    <p:input port="xslt">
      <p:pipe step="docx2html" port="docx2hub-xslt"/>
    </p:input>
  </docx2hub:convert>

  <hub2htm:convert name="hub2html">
    <p:input port="paths"><p:empty/></p:input>    
    <p:input port="other-params"><p:empty/></p:input>    
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
  </hub2htm:convert>

</p:declare-step>