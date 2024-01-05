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
  name="apply-changemarkup"
  type="docx2hub:apply-changemarkup">

  <p:input port="source" primary="true">
    <p:documentation>The result of docx2hub:single-tree (or mathtype2mml)</p:documentation>
  </p:input>
  <p:input port="params">
    <p:documentation>The params output port of docx2hub:single-tree</p:documentation>
  </p:input>
  <p:input port="stylesheet">
    <p:document href="../xsl/main.xsl"/>
  </p:input>
  <p:input port="schematron">
    <p:document href="../sch/changemarkup.sch.xml"/>
  </p:input>

  <p:output port="result" primary="true">
    <p:documentation>The same basic structure as the primary source of the current step, but with applied markup changes (such as text deletions or image insertions)</p:documentation>
    <p:pipe port="result" step="convert-apply-changemarkup"/>
  </p:output>
  <p:output port="report" sequence="true">
    <p:pipe port="report" step="convert-apply-changemarkup"/>
  </p:output>
  <p:output port="schema" sequence="true">
    <p:pipe port="result" step="schematron-atts"/>
  </p:output>

  <p:serialization port="result" omit-xml-declaration="false"/>
  
  <p:option name="debug" required="false" select="'no'"/>
  <p:option name="debug-dir-uri" required="false" select="'file:/tmp/debug'"/>
  <p:option name="fail-on-error" select="'no'"/>
  <p:option name="active" required="false" select="'yes'">
    <p:documentation>Whether apply-changemarkup conversion happens at all.</p:documentation>
  </p:option>
  
  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>
  <p:import href="http://transpect.io/xproc-util/xslt-mode/xpl/xslt-mode.xpl"/>

  <p:choose name="convert-apply-changemarkup">
    <p:when test="$active = 'yes' and exists(//w:del | //w:moveFrom | //w:ins)">
      <p:output port="result" primary="true">
        <p:pipe port="result" step="apply-changemarkup-xslt"/>
      </p:output>
      <p:output port="report" sequence="true">
        <p:pipe port="result" step="check"/>
      </p:output>
      
      <p:variable name="basename" select="/c:param-set/c:param[@name = 'basename']/@value">
        <p:pipe step="apply-changemarkup" port="params"/>
      </p:variable>
      
      <tr:xslt-mode msg="yes" mode="docx2hub:apply-changemarkup" name="apply-changemarkup-xslt">
        <p:input port="parameters">
          <p:pipe step="apply-changemarkup" port="params"/>
        </p:input>
        <p:input port="stylesheet">
          <p:pipe step="apply-changemarkup" port="stylesheet"/>
        </p:input>
        <p:input port="models">
          <p:empty/>
        </p:input>
        <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/02a-apply-changemarkup')"/>
        <p:with-option name="debug" select="$debug"/>
        <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
        <p:with-option name="fail-on-error" select="$fail-on-error"/>
        <p:with-param name="fail-on-error" select="$fail-on-error"/>
      </tr:xslt-mode>
      
      <p:validate-with-schematron assert-valid="false" name="val-sch">
        <p:input port="schema">
          <p:pipe port="schematron" step="apply-changemarkup"/>
        </p:input>
        <p:input port="parameters"><p:empty/></p:input>
        <p:with-param name="allow-foreign" select="'true'"/>
      </p:validate-with-schematron>
    
      <p:sink/>
    
      <p:add-attribute name="check0" match="/*" 
        attribute-name="tr:rule-family" attribute-value="docx2hub_changemarkup">
        <p:input port="source">
          <p:pipe port="report" step="val-sch"/>
        </p:input>
      </p:add-attribute>
      
      <p:insert name="check" match="/*" position="first-child">
        <p:input port="insertion" select="/*/*:title">
          <p:pipe port="schematron" step="apply-changemarkup"/>
        </p:input>
      </p:insert>
      
    </p:when>
    <p:when test="$active = 'check'">
      <p:output port="result" primary="true">
        <p:pipe port="result" step="apply-changemarkup-identity"/>
      </p:output>
      <p:output port="report" sequence="true">
        <p:pipe port="result" step="check"/>
      </p:output>
      
      <p:identity name="apply-changemarkup-identity"/>
      
      <p:validate-with-schematron assert-valid="false" name="val-sch">
        <p:input port="schema">
          <p:pipe port="schematron" step="apply-changemarkup"/>
        </p:input>
        <p:input port="parameters"><p:empty/></p:input>
        <p:with-param name="allow-foreign" select="'true'"/>
      </p:validate-with-schematron>
      
      <p:sink/>
      
      <p:add-attribute name="check0" match="/*" 
        attribute-name="tr:rule-family" attribute-value="docx2hub_changemarkup">
        <p:input port="source">
          <p:pipe port="report" step="val-sch"/>
        </p:input>
      </p:add-attribute>
      
      <p:insert name="check" match="/*" position="first-child">
        <p:input port="insertion" select="/*/*:title">
          <p:pipe port="schematron" step="apply-changemarkup"/>
        </p:input>
      </p:insert>
      
    </p:when>
    <p:otherwise>
      <p:output port="result" primary="true"/>
      <p:output port="report" sequence="true">
        <p:empty/>
      </p:output>
      <p:identity/>
    </p:otherwise>
  </p:choose>

  <p:sink/>
  
  <p:add-attribute match="/*" 
    attribute-name="tr:step-name" 
    attribute-value="docx2hub">
    <p:input port="source">
      <p:pipe port="schematron" step="apply-changemarkup"/>
    </p:input>
  </p:add-attribute>
  
  <p:add-attribute 
    match="/*" 
    attribute-name="tr:rule-family" 
    attribute-value="docx2hub">
  </p:add-attribute>
  
  <p:insert name="schematron-atts" match="/*" position="first-child">
    <p:input port="insertion" select="/*/*:title">
      <p:pipe port="schematron" step="apply-changemarkup"/>
    </p:input>
  </p:insert>
  
  <p:sink/>
  
</p:declare-step>
