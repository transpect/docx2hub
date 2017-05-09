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

  <p:output port="result" primary="true">
    <p:documentation>The same basic structure as the single tree, but with equation OLE objects replaced with MathML</p:documentation>
  </p:output>
  <p:output port="modified-zip-manifest">
    <p:pipe port="result" step="zip-manifest"/>
  </p:output>
  <p:output port="report" sequence="true">
    <p:pipe port="report" step="check"/>
  </p:output>

  <p:serialization port="result" omit-xml-declaration="false"/>
  
  <p:option name="debug" required="false" select="'no'"/>
  <p:option name="debug-dir-uri" required="false" select="'file:/tmp/debug'"/>
  <p:option name="active" required="false" select="'yes'">
    <p:documentation>Whether MathType→MathML conversion happens at all.</p:documentation>
  </p:option>
  <p:option name="mml-space-handling" select="'mspace'">
    <p:documentation>see corresponding documentation for docx2hub</p:documentation>
  </p:option>
  
  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>
  <p:import href="http://transpect.io/calabash-extensions/mathtype-extension/xpl/mathtype2mml-declaration.xpl"/>
  <p:import href="http://transpect.io/xproc-util/store-debug/xpl/store-debug.xpl"/>


  <p:choose name="convert-mathtype2mml">
    <p:when test="$active eq 'yes'">
      <p:variable name="basename" select="/c:param-set/c:param[@name = 'basename']">
        <p:pipe port="params" step="mathtype2mml"/>
      </p:variable>
      <p:viewport match="/w:root/w:document/w:body/*[local-name() = ('p', 'tbl')]
                           //w:object/o:OLEObject[@Type eq 'Embed' and starts-with(@ProgID, 'Equation')]" name="mathtype2mml-viewport">
        <p:variable name="rel-id" select="o:OLEObject/@r:id"/>
        <p:variable name="equation-href" select="concat(/w:root/@xml:base,
                                                        'word/',
                                                        /w:root/w:docRels/rel:Relationships/rel:Relationship[@Id eq $rel-id]/@Target
                                               )">
          <p:pipe port="result" step="insert-xpath"/>
        </p:variable>
        
        <p:try name="mathtype2mml-wrapper">
          <p:group>
            <tr:mathtype2mml name="mathtype2mml">
              <p:with-option name="href" select="$equation-href"/>
              <p:with-option name="debug" select="$debug"/>
              <p:with-option name="debug-dir-uri" select="concat($debug-dir-uri, '/docx2hub/', $basename, '/')"/>
            </tr:mathtype2mml>
          </p:group>
          <p:catch>
            <p:identity/>
          </p:catch>
        </p:try>
         
      </p:viewport>
           
      <tr:store-debug>
        <p:with-option name="pipeline-step" select="concat('docx2hub/', $basename, '/02b-mathtype-converted')"/>
        <p:with-option name="active" select="$debug"/>
        <p:with-option name="base-uri" select="$debug-dir-uri"/>
      </tr:store-debug>
    </p:when>
    <p:otherwise>
      <p:identity/>
    </p:otherwise>
  </p:choose>

  <p:add-attribute attribute-name="mathtype2mml" match="/*" name="add-mathtype2mml-attr">
    <p:with-option name="attribute-value" select="$mathtype2mml"/>
  </p:add-attribute>
  

</p:declare-step>
