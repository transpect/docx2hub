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
  name="single-tree-enhanced"
  type="docx2hub:single-tree-enhanced">

  <p:documentation>For usage in docx_modify and other XProc steps with need of the docx single-tree plus some more enhancements.</p:documentation>
  
  <p:input port="xslt">
    <p:document href="../xsl/main.xsl"/>
  </p:input>
  <p:input port="source" primary="true">
    <p:documentation>This is to prevent a default readable port connecting to this step’s xslt port.</p:documentation>
    <p:empty/>
  </p:input>
  <p:input port="single-tree-schematron">
    <p:document href="../sch/single-tree.sch.xml"/>
    <p:documentation>Schematron that will validate the entire Word container document.</p:documentation>
  </p:input>
  <p:input port="change-markup-schematron">
    <p:document href="../sch/changemarkup.sch.xml"/>
    <p:documentation>Schematron that will validate the entire document after applying change markup.</p:documentation>
  </p:input>
  <p:input port="mathtype2mml-schematron">
    <p:document href="../sch/mathtype2mml.sch.xml"/>
    <p:documentation>Schematron that will validate the entire document after replacing MathType OLE-Objects by MathML.</p:documentation>
  </p:input>
  <p:input port="custom-font-maps" primary="false" sequence="true">
    <p:documentation>
      See additional-font-maps in mathtype-extension
    </p:documentation>
    <p:empty/>
  </p:input>
  
  <p:output port="result" primary="true"/>
  <p:output port="report" sequence="true">
    <p:pipe port="report" step="single-tree"/>
    <p:pipe port="report" step="apply-changemarkup"/>
    <p:pipe port="report" step="mathtype2mml"/>
  </p:output>
  <p:output port="zip-manifest">
    <p:pipe port="modified-zip-manifest" step="mathtype2mml"/>
  </p:output>
  <p:output port="params">
    <p:pipe port="params" step="single-tree"/>
  </p:output>
  <p:output port="schema" sequence="true">
    <p:pipe port="schema" step="single-tree"/>
    <p:pipe port="schema" step="apply-changemarkup"/>
    <p:pipe port="schema" step="mathtype2mml"/>
  </p:output>
  
  <!-- Options: See documentation in docx2hub.xpl-->
  <p:option name="docx" required="true"/>
  <p:option name="debug" select="'no'"/>
  <p:option name="debug-dir-uri" select="'debug'"/>
  <p:option name="status-dir-uri" select="'status'"/>
  <p:option name="srcpaths" select="'no'"/>
  <p:option name="no-srcpaths-for-text-runs-threshold" select="'40000'"/>
  <p:option name="unwrap-tooltip-links" select="'no'"/>
  <p:option name="mml-space-handling" select="'mspace'"/>
  <p:option name="hub-version" select="'1.2'"/>
  <p:option name="fail-on-error" select="'no'"/>
  <p:option name="field-vars" select="'no'"/>
  <p:option name="extract-dir" select="''"/>
  <p:option name="mathtype2mml" required="false" select="'yes'"/>
  <p:option name="mathtype-source-pi" required="false" select="'no'"/>
  <p:option name="mathtype2mml-cleanup" required="false" select="'yes'"/>
  <p:option name="apply-changemarkup" required="false" select="'yes'"/>
  <p:option name="use-filename-from-http-response" required="false" select="'no'"/>

  <p:import href="single-tree.xpl"/>
  <p:import href="apply-changemarkup.xpl"/>
  <p:import href="mathtype2mml.xpl"/>

  <docx2hub:single-tree name="single-tree">
    <p:with-option name="docx" select="$docx"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-option name="unwrap-tooltip-links" select="$unwrap-tooltip-links"/>
    <p:with-option name="mml-space-handling" select="$mml-space-handling"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-option name="field-vars" select="$field-vars"/>
    <p:with-option name="srcpaths" select="$srcpaths"/>
    <p:with-option name="no-srcpaths-for-text-runs-threshold" select="$no-srcpaths-for-text-runs-threshold"/>
    <p:with-option name="extract-dir" select="$extract-dir"/>
    <p:with-option name="use-filename-from-http-response" select="$use-filename-from-http-response"/>
    <p:input port="schematron">
      <p:pipe step="single-tree-enhanced" port="single-tree-schematron"/>
    </p:input>
    <p:input port="xslt">
      <p:pipe step="single-tree-enhanced" port="xslt"/>
    </p:input>
  </docx2hub:single-tree>
  
  <docx2hub:apply-changemarkup name="apply-changemarkup">
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="active" select="$apply-changemarkup"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:input port="params">
      <p:pipe step="single-tree" port="params"/>
    </p:input>
    <p:input port="schematron">
      <p:pipe step="single-tree-enhanced" port="change-markup-schematron"/>
    </p:input>
  </docx2hub:apply-changemarkup>

  <docx2hub:mathtype2mml name="mathtype2mml">
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="mml-space-handling" select="$mml-space-handling"/>
    <p:with-option name="active" select="$mathtype2mml"/>
    <p:with-option name="word-container-cleanup" select="$mathtype2mml-cleanup"/>
    <p:with-option name="sources" select="$mathtype2mml"/>
    <p:with-option name="mathtype-source-pi" select="$mathtype-source-pi"/>
    <p:input port="params">
      <p:pipe step="single-tree" port="params"/>
    </p:input>
    <p:input port="schematron">
      <p:pipe step="single-tree-enhanced" port="mathtype2mml-schematron"/>
    </p:input>
    <p:input port="zip-manifest">
      <p:pipe step="single-tree" port="zip-manifest"/>
    </p:input>
    <p:input port="custom-font-maps">
      <p:pipe port="custom-font-maps" step="single-tree-enhanced"/>
    </p:input>
  </docx2hub:mathtype2mml>

</p:declare-step>
