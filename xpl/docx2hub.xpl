<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" 
  xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions" 
  xmlns:tr="http://transpect.io"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
  version="1.0" 
  name="docx2hub" 
  type="docx2hub:convert">

  <p:documentation xmlns="http://www.w3.org/1999/xhtml">
    <p>This script is used to convert docx to Hub XML. By default, the output is stored in the same directory as the input docx
      file, with the same basename. It is a library that needs other externals. For standalone operation, please check out the
        <a href="https://subversion.le-tex.de/docxtools/trunk/docx2hub_frontend/">front-end project</a>.</p>
    <p>From the directory wherey you checked out the front-end project, you invoke it with:</p>
    <p><code>calabash/calabash.sh docx2hub/wml2hub.xpl docx=PATH-TO-MY-DOCX-FILE.docx</code></p>
    <p>where docx may be an OS path or a file:, http:, or https: URL.</p>
    <p>Import it with</p>
    <p><code>&lt;p:import href="http://transpect.io/docx2hub/xpl/wml2hub.xpl" /></code></p>
    <p>if you use it from transpect or if you imported this project as svn:external.</p>
    <p>In the latter case, include the following line in you project's xmlcatalog/catalog.xml:</p>
    <p><code>&lt;nextCatalog catalog="../docx2hub/xmlcatalog/catalog.xml"/></code></p>
    <p>Experts may override the default conversion rules by supplying custom XSLT (that imports main.xsl) on the 'xslt'
      port.</p>
  </p:documentation>

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
  <p:input port="changemarkup-schematron">
    <p:document href="../sch/changemarkup.sch.xml"/>
    <p:documentation>Schematron that will validate the entire document after applying change markup.</p:documentation>
  </p:input>
  <p:input port="mathtype2mml-schematron">
    <p:document href="../sch/mathtype2mml.sch.xml"/>
    <p:documentation>Schematron that will validate the entire document after replacing MathType OLE-Objects by MathML.</p:documentation>
  </p:input>
  <p:input port="field-functions-schematron">
    <p:document href="../sch/field-functions.sch.xml"/>
    <p:documentation>Schematron that will validate the intermediate format after merging/splitting Word field functions.</p:documentation>
  </p:input>
  <p:input port="result-schematron">
    <p:document href="../sch/result.sch.xml"/>
    <p:documentation>Schematron that will validate the flat Hub. It will chiefly report error messages that were 
      embedded during conversion.</p:documentation>
  </p:input>
  <p:input port="custom-font-maps" primary="false" sequence="true">
    <p:documentation>See same port in mathtype2mml.xpl.
    If you need to match a specific docx font name that is not identical with the base name of the base URI’s file name part,
    you can do so by including an attribute /symbols/@docx-name.</p:documentation>
    <p:empty/>
  </p:input>
  <p:output port="result" primary="true"/>
  <p:serialization port="result" omit-xml-declaration="false"/>
  <p:output port="insert-xpath">
    <p:pipe step="single-tree-enhanced" port="result"/>
  </p:output>
  <p:output port="report" sequence="true">
    <p:pipe port="report" step="single-tree-enhanced"/>
    <p:pipe port="report" step="add-props"/>
    <p:pipe port="report" step="props2atts"/>
    <p:pipe port="report" step="remove-redundant-run-atts"/>
    <p:pipe port="report" step="join-instrText-runs"/>
    <p:pipe port="report" step="field-functions"/>
    <p:pipe port="result" step="check-field-functions"/>
    <p:pipe port="report" step="wml-to-dbk"/>
    <p:pipe port="report" step="join-runs"/>
    <p:pipe port="result" step="check-result"/>
    <p:pipe port="result" step="rename-errorPI2svrl-reports"/>
    <p:pipe port="report" step="check-tables"/>
  </p:output>
  <p:output port="schema" sequence="true">
    <p:pipe port="result" step="decorate-field-functions-schematron"/>
    <p:pipe port="result" step="decorate-result-schematron"/>
    <p:pipe port="schema" step="single-tree-enhanced"/>
  </p:output>
  <p:output port="zip-manifest">
    <p:pipe port="zip-manifest" step="single-tree-enhanced"/>
  </p:output>



  <p:option name="docx" required="true">
    <p:documentation>OS name (preferably with full path, may not resolve if only a relative path is given), file:, http:, or
      https: URL. The file will be fetched first if it is an HTTP URL.</p:documentation>
  </p:option>
  <p:option name="debug" select="'no'"/>
  <p:option name="debug-dir-uri" select="'debug'"/>
  <p:option name="status-dir-uri" select="'status'"/>
  <p:option name="srcpaths" select="'no'"/>
  <p:option name="unwrap-tooltip-links" select="'no'"/>
  <p:option name="mml-space-handling" select="'mspace'">
    <p:documentation xmlns="http://www.w3.org/1999/xhtml">
      <p>Whitespace conversion from OMML to MathML</p>
      <dl>
        <dt>none</dt>
        <dd>All whitespace will be in mtext, without xml:space attributes</dd>
        <dt>xml-space</dt>
        <dd>Same mtext as for 'none', with xml:space="preserve" attribute when \s at the beginning or end</dd>
        <dt>figure-space</dt>
        <dd>Each \s char will converted to U+2007, still within mtext. Special handling for tabs and newlines tbd</dd>
        <dt>mspace</dt>
        <dd>Start and end \s will be converted to mspace with a width of 0.25em for each char</dd>
      </dl>
    </p:documentation>
  </p:option>
  <p:option name="hub-version" select="'1.2'"/>
  <p:option name="fail-on-error" select="'no'"/>
  <p:option name="field-vars" select="'no'"/>
  <p:option name="extract-dir" select="''">
    <p:documentation>Directory (OS path, not file: URL) to which the file will be unzipped. If option is empty string, will be
      '.tmp' appended to OS file path.
    Will be available as extract-dir-uri in the c:param-set document that comes out of single-tree-enhanced on the
    params port.</p:documentation>
  </p:option>
  <p:option name="create-svg" required="false" select="'no'">
    <p:documentation>Whether Office Open Drawing ML should be mapped to SVG</p:documentation>
  </p:option>
  <p:option name="discard-alternate-choices" select="'yes'">
    <p:documentation>Whether to remove mc:AlternateContent/mc:Choice at an early conversion stage (after insert-xpath 
      though).</p:documentation>
  </p:option>
  <p:option name="charmap-policy" select="'unicode'">
    <p:documentation>Pass a policy for mapping characters from non-unicode fonts: 
      'unicode' maps characters strictly to their unicode equivalents. Please be aware that 
      preinstalled fonts may not be able to display uncommon characters correctly.
      'msoffice' tries to map each character in a way that it can be displayed with typical MS Office fonts, 
      even if the appearance of the character doesn't match exactly those of its source. 
    </p:documentation>
  </p:option>
  <p:option name="mathtype2mml" required="false" select="'yes'">
    <p:documentation xmlns="http://www.w3.org/1999/xhtml">
      <p>Activates use of mathtype2mml extension.</p>
      <p>Should be one of the following String values:</p>
      <dl>
        <dt>no</dt>
          <dd>no conversion happens</dd>
        <dt>ole</dt>
          <dd>Use the OLE-Object as source for equation.</dd>
        <dt>wmf</dt>
          <dd>
            Use the wmf-image as source for MTEF equation. <br/>
            If no equation is found in wmf file, OLE-Object is used as fallback
          </dd>
        <dt>'wmf+ole' | 'ole+wmf'</dt>
          <dd>
            Use both, wmf and OLE as source for MTEF equation <br/>
            If both are deep-equal, only one MathML equation will be output. <br/>
            If they differ, both equations will be output. <br/>
            In addition, a processing-instruction will be added after the MathML equation, stating that the MathML equations (from wmf and ole sources) differ. <br/>
            The order is defined by the order in the String (wmf+ole makes wmf equation appear first, ole+wmf makes ole equation appear first).
        <dt>yes</dt>
          <dd>
            Same as 'ole'. <br/>
            This is the default value for the option.
          </dd>
        <dt>any other String</dt>
          <dd>Is treated as 'yes'.</dd>
          </dd>
      </dl>
    </p:documentation>
  </p:option>
  <p:option name="mathtype-source-pi" required="false" select="'no'">
    <p:documentation xmlns="http://www.w3.org/1999/xhtml">
      <p>If not set to 'no', each mml-equation also bears a processing-instruction stating its source (file ending):
        <dl>
          <dt>M2M_210</dt>
          <dd>MathML equation source:ole</dd>
          <dt>M2M_211</dt>
          <dd>MathML equation source:wmf</dd>
        </dl>
      </p>
    </p:documentation>
  </p:option>
  <p:option name="apply-changemarkup" required="false" select="'yes'">
    <p:documentation xmlns="http://www.w3.org/1999/xhtml">
      <p>Apply all change markup on the compound word document.</p>
    </p:documentation>
  </p:option>
  <p:option name="use-filename-from-http-response" required="false" select="'no'">
    <p:documentation>Use filename that is passed on from http request response instead of 
    possible filename read from URL in tr:file-uri (for example when using Gdocs URLs:
    https://docs.google.com/document/d/1Z5eYyjLoRhB24HYZ-d-wQKAFD3QDWZUsQH4cKHs2eiM/export?format=docx)</p:documentation>
  </p:option>
  <p:option name="check-tables" required="false" select="'no'">
    <p:documentation>If this option is set to 'yes' all tables are normalized with calstable and checked against 
    schematron</p:documentation>
  </p:option>
  <p:option name="include-header-and-footer" required="false" select="'no'">
    <p:documentation>Whether to include headers and footers as div at the beginning of the document. Permitted values: yes|no</p:documentation>
  </p:option>

  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>
  <p:import href="http://transpect.io/calabash-extensions/unzip-extension/unzip-declaration.xpl"/>
  <p:import href="http://transpect.io/xproc-util/file-uri/xpl/file-uri.xpl"/>
  <p:import href="http://transpect.io/xproc-util/xml-model/xpl/prepend-hub-xml-model.xpl"/>
  <p:import href="http://transpect.io/xproc-util/xslt-mode/xpl/xslt-mode.xpl"/>
  <p:import href="http://transpect.io/xproc-util/store-debug/xpl/store-debug.xpl"/>
  <p:import href="http://transpect.io/xproc-util/simple-progress-msg/xpl/simple-progress-msg.xpl" 
    use-when="doc-available('http://transpect.io/xproc-util/simple-progress-msg/xpl/simple-progress-msg.xpl')"/>
  <p:import href="single-tree-enhanced.xpl"/>
  <p:import href="mathtype2mml.xpl"/>
  <p:import href="apply-changemarkup.xpl"/>
  <p:import href="http://transpect.io/htmlreports/xpl/errorPI2svrl.xpl"/>

  <p:group use-when="doc-available('http://transpect.io/xproc-util/simple-progress-msg/xpl/simple-progress-msg.xpl')">
    <tr:simple-progress-msg name="start-msg" file="docx2hub-start.txt">
      <p:input port="msgs">
        <p:inline>
          <c:messages>
            <c:message xml:lang="en">Starting DOCX to flat Hub XML conversion</c:message>
            <c:message xml:lang="de">Beginne Konvertierung von DOCX zu flachem Hub XML</c:message>
          </c:messages>
        </p:inline>
      </p:input>
      <p:with-option name="status-dir-uri" select="$status-dir-uri"/>
    </tr:simple-progress-msg>
  </p:group>

  <docx2hub:single-tree-enhanced name="single-tree-enhanced">
    <p:with-option name="docx" select="$docx"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="status-dir-uri" select="$status-dir-uri"/>
    <p:with-option name="srcpaths" select="$srcpaths"/>
    <p:with-option name="unwrap-tooltip-links" select="$unwrap-tooltip-links"/>
    <p:with-option name="mml-space-handling" select="$mml-space-handling"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-option name="field-vars" select="$field-vars"/>
    <p:with-option name="extract-dir" select="$extract-dir"/>
    <p:with-option name="mathtype2mml" select="$mathtype2mml"/>
    <p:with-option name="mathtype-source-pi" select="$mathtype-source-pi"/>
    <p:with-option name="use-filename-from-http-response" select="$use-filename-from-http-response"/>
    <p:with-option name="apply-changemarkup" select="$apply-changemarkup"/>
    <p:input port="single-tree-schematron">
      <p:pipe step="docx2hub" port="single-tree-schematron"/>
    </p:input>
    <p:input port="change-markup-schematron">
      <p:pipe step="docx2hub" port="changemarkup-schematron"/>
    </p:input>
    <p:input port="mathtype2mml-schematron">
      <p:pipe step="docx2hub" port="mathtype2mml-schematron"/>
    </p:input>
    <p:input port="xslt">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="custom-font-maps">
      <p:pipe port="custom-font-maps" step="docx2hub"/>
    </p:input>
  </docx2hub:single-tree-enhanced>
  
  <tr:xslt-mode msg="yes" mode="docx2hub:preprocess-styles" name="preprocess-styles">
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/03a')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="field-vars" select="$field-vars"/>
    <p:with-param name="mathtype2mml" select="$mathtype2mml"/>
  </tr:xslt-mode>
  
  <tr:xslt-mode msg="yes" mode="docx2hub:resolve-tblBorders" name="resolve-tblBorders">
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/03b')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="field-vars" select="$field-vars"/>
    <p:with-param name="mathtype2mml" select="$mathtype2mml"/>
  </tr:xslt-mode>
  
  <tr:xslt-mode msg="yes" mode="docx2hub:add-props" name="add-props">
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/04')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="field-vars" select="$field-vars"/>
    <p:with-param name="mathtype2mml" select="$mathtype2mml"/>
    <p:with-param name="discard-alternate-choices" select="$discard-alternate-choices"/>
    <p:with-param name="include-header-and-footer" select="$include-header-and-footer"/>
  </tr:xslt-mode>

  <tr:xslt-mode msg="yes" mode="docx2hub:props2atts" name="props2atts">
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/05')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </tr:xslt-mode>

  <tr:xslt-mode msg="yes" mode="docx2hub:remove-redundant-run-atts" name="remove-redundant-run-atts">
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/07')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </tr:xslt-mode>

  <p:sink/>

  <tr:xslt-mode msg="yes" mode="docx2hub:join-instrText-runs" name="join-instrText-runs">
    <p:input port="source">
      <p:pipe port="result" step="remove-redundant-run-atts"/>
      <p:pipe port="custom-font-maps" step="docx2hub"/>
      <p:document href="http://this.transpect.io/xmlcatalog/catalog.xml"/>
    </p:input>
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/08')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </tr:xslt-mode>

  <p:viewport match="w:instrText[starts-with(@docx2hub:field-function-name, 'CITAVI')]" name="citavi-viewport">
    <p:documentation>Since there seems to be no other means in XProc 1.0 (with Saxon HE that does not support the
      EXPath binary module) to decode base64, we will store these contents to JSON files and read them later using 
      XPath 3.1 json-doc(). The same for CITAVI.BIBLIOGRAPHY which apparently is base64-encoded XML.</p:documentation>
    <p:output port="result" primary="true">
      <p:pipe port="result" step="new-Citavi-field-code"/>
    </p:output>
    <p:variable name="pos" select="p:iteration-position()"/>
    <p:variable name="filetype" select="lower-case(replace(/*/@docx2hub:field-function-name, 'CITAVI_', ''))"/>
    <p:variable name="filename" 
      select="concat(/c:param-set/c:param[@name='extract-dir-uri']/@value, 'Citavi/Citavi_', $pos, '.', $filetype)">
      <p:pipe port="params" step="single-tree-enhanced"/>
    </p:variable>
    <p:add-attribute attribute-name="docx2hub:field-function-args" match="/*" name="new-Citavi-field-code">
      <p:with-option name="attribute-value" select="$filename"/>
    </p:add-attribute>
    <p:sink name="citavi-sink1"/>
    <p:add-attribute match="/*" name="add-citavi-content-type" attribute-name="content-type">
      <p:input port="source">
        <p:inline><c:data encoding="base64">dummy</c:data></p:inline>
      </p:input>      
      <p:with-option name="attribute-value" select="concat('application/', $filetype)"/>
    </p:add-attribute>
    <p:string-replace match="/c:data/text()" name="citavi-base64-to-c_data">
      <p:with-option name="replace" select="concat('''', /*/@docx2hub:field-function-args, '''')">
        <p:pipe port="current" step="citavi-viewport"/>
      </p:with-option>
    </p:string-replace>
    <p:store cx:decode="true" name="store-citavi-file">
      <p:with-option name="href" select="$filename"/>
    </p:store>
  </p:viewport>
  
  <p:choose>
    <p:when test="exists(.//w:instrText[starts-with(@docx2hub:field-function-name, 'CITAVI')])">
      <tr:store-debug>
        <p:with-option name="pipeline-step"
          select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/09-citavi-viewport')">
          <p:pipe port="params" step="single-tree-enhanced"/>
        </p:with-option>
        <p:with-option name="active" select="$debug"/>
        <p:with-option name="base-uri" select="$debug-dir-uri"/>
      </tr:store-debug>
    </p:when>
    <p:otherwise>
      <p:identity/>
    </p:otherwise>
  </p:choose>

  <tr:xslt-mode msg="yes" mode="docx2hub:field-functions" name="field-functions">
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/14')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </tr:xslt-mode>

  <p:validate-with-schematron assert-valid="false" name="check-field-functions0">
    <p:input port="schema">
      <p:pipe port="field-functions-schematron" step="docx2hub"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
    <p:with-param name="allow-foreign" select="'true'"/>
  </p:validate-with-schematron>

  <p:sink/>

  <p:add-attribute match="/*" 
    attribute-name="tr:step-name" attribute-value="docx2hub">
    <p:input port="source">
      <p:pipe port="report" step="check-field-functions0"/>
    </p:input>
  </p:add-attribute>
  
  <p:add-attribute name="check-field-functions1" match="/*" 
    attribute-name="tr:rule-family" attribute-value="docx2hub_field-functions">
    <p:documentation>Will also check other things such as change markup.</p:documentation>
  </p:add-attribute>
  
  <p:insert name="check-field-functions" match="/*" position="first-child">
    <p:input port="insertion" select="/*/*:title">
      <p:pipe port="field-functions-schematron" step="docx2hub"/>
    </p:input>
  </p:insert>
  
  <p:sink/>
  
  <p:add-attribute match="/*" 
                   attribute-name="tr:step-name" 
                   attribute-value="docx2hub">
    <p:input port="source">
      <p:pipe port="field-functions-schematron" step="docx2hub"/>
    </p:input>
  </p:add-attribute>
  
  <p:add-attribute name="decorate-field-functions-schematron0" 
                   match="/*" 
                   attribute-name="tr:rule-family" 
                   attribute-value="docx2hub">
  </p:add-attribute>
  
  <p:insert name="decorate-field-functions-schematron" match="/*" position="first-child">
    <p:input port="insertion" select="/*/*:title">
      <p:pipe port="field-functions-schematron" step="docx2hub"/>
    </p:input>
  </p:insert>
  
  <p:sink/>

  <tr:xslt-mode msg="yes" mode="wml-to-dbk" name="wml-to-dbk">
    <p:input port="source">
      <p:pipe port="result" step="field-functions"/>
      <p:pipe port="custom-font-maps" step="docx2hub"/>
      <p:document href="http://this.transpect.io/xmlcatalog/catalog.xml"/>
    </p:input>
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/20')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-param name="srcpaths" select="$srcpaths"/>
    <p:with-param name="unwrap-tooltip-links" select="$unwrap-tooltip-links"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="field-vars" select="$field-vars"/>
    <p:with-param name="charmap-policy" select="$charmap-policy"/>
  </tr:xslt-mode>

  <p:sink/>

  <tr:xslt-mode msg="yes" mode="docx2hub:join-runs" name="join-runs">
    <p:input port="source">
      <p:pipe port="result" step="wml-to-dbk"/>
      <p:document href="http://this.transpect.io/xmlcatalog/catalog.xml"/>
    </p:input>
    <p:input port="parameters">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[@name='basename']/@value, '/24')">
      <p:pipe port="params" step="single-tree-enhanced"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </tr:xslt-mode>

  <p:add-attribute match="/*" attribute-name="xml:base" name="rebase">
    <p:with-option name="attribute-value" 
      select="replace(/c:param-set/c:param[@name='local-href']/@value, '\.do[ct][xm]$', '.hub.xml')">
      <p:pipe step="single-tree-enhanced" port="params"/>
    </p:with-option>
  </p:add-attribute>

  <tr:errorPI2svrl name="errorPI2svrl" pi-names="tr" group-by-srcpath="no">
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="status-dir-uri" select="$status-dir-uri"/>
  </tr:errorPI2svrl>
  
  <p:sink/>
  
  <!--  table normalization -->
  <p:choose name="check-tables">
    <p:when test="$check-tables='yes'">
      <p:output port="normalized-tables">
        <p:pipe port="result" step="normalize-tables"/>
      </p:output>
      
      <p:output port="report">
        <p:pipe port="result" step="sch_tables"/>
      </p:output>
      
      <p:xslt name="normalize-tables">
        <p:input port="source">
          <p:pipe port="result" step="join-runs"/>
        </p:input>
        <p:input port="stylesheet">
          <p:inline>
            <xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
              xmlns:calstable="http://docs.oasis-open.org/ns/oasis-exchange/table"
              version="2.0">
              
              <xsl:import href="http://transpect.io/xslt-util/calstable/xsl/normalize.xsl"/>
              
              <xsl:template match="node() | @*" >
                <xsl:copy>
                  <xsl:apply-templates select="@*, node()"/>
                </xsl:copy>
              </xsl:template>
              
              <xsl:template match="*[*:row]">
                <xsl:sequence select="calstable:check-normalized(
                                          calstable:normalize(.), 
                                          'no'
                                          )"/>
              </xsl:template>
              
            </xsl:stylesheet>
          </p:inline>
        </p:input>
        <p:input port="parameters">
          <p:pipe step="single-tree-enhanced" port="params"/>
        </p:input>
      </p:xslt>
      
      <p:validate-with-schematron assert-valid="false" name="sch_tables0">
        <p:input port="schema">
          <p:document href="http://transpect.io/xslt-util/calstable/sch/sch_tables.sch.xml"/>
        </p:input>
        <p:input port="parameters"><p:empty/></p:input>
        <p:with-param name="allow-foreign" select="'true'"/>
      </p:validate-with-schematron>
      
      <p:sink/>
      
      <p:add-attribute match="/*" 
        attribute-name="tr:step-name" attribute-value="docx2hub">
        <p:input port="source">
          <p:pipe port="report" step="sch_tables0"/>
        </p:input>
      </p:add-attribute>
      
      <p:add-attribute name="sch_tables1" match="/*" 
        attribute-name="tr:rule-family" attribute-value="docx2hub_tables">
      </p:add-attribute>
      
      <p:insert name="sch_tables" match="/*" position="first-child">
        <p:input port="insertion" select="/*/*:title">
          <p:document href="http://transpect.io/xslt-util/calstable/sch/sch_tables.sch.xml"/>
        </p:input>
      </p:insert>
      
      
    </p:when>
    
    <p:otherwise>
        <p:output port="normalized-tables">
        <p:pipe port="result" step="identity"/>
      </p:output>
      
      <p:output port="report">
        <p:empty/>
      </p:output>
      
      <p:identity name="identity">
        <p:input port="source">
          <p:pipe port="result" step="join-runs"/>
        </p:input>
      </p:identity>
      
    </p:otherwise>
  </p:choose>
  
  <p:for-each name="rename-errorPI2svrl-reports">
    <p:output port="result" primary="true"/>
    <p:iteration-source>
      <p:pipe port="report" step="errorPI2svrl"/>
    </p:iteration-source>
    <p:add-attribute attribute-name="tr:rule-family" match="/*" name="rename-errorPI2svrl-reports0">
      <p:with-option name="attribute-value" select="replace(/*/@tr:rule-family, '^W2D', 'docx2hub_PI')"/>
    </p:add-attribute>
  </p:for-each>
  
  <p:sink/>

  <tr:prepend-hub-xml-model name="pi">
    <p:input port="source">
      <p:pipe port="result" step="errorPI2svrl"/>
    </p:input>
    <p:with-option name="hub-version" select="$hub-version"/>
  </tr:prepend-hub-xml-model>

  <p:validate-with-schematron assert-valid="false" name="check-result0">
    <p:input port="schema">
      <p:pipe port="result-schematron" step="docx2hub"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
    <p:with-param name="allow-foreign" select="'true'"/>
  </p:validate-with-schematron>

  <p:sink/>

  <p:add-attribute name="check-result1" match="/*" 
    attribute-name="tr:rule-family" attribute-value="docx2hub_result">
    <p:input port="source">
      <p:pipe port="report" step="check-result0"/>
    </p:input>
  </p:add-attribute>
  
  <p:insert name="check-result" match="/*" position="first-child">
    <p:input port="insertion" select="/*/*:title">
      <p:pipe port="result-schematron" step="docx2hub"/>
    </p:input>
  </p:insert>
  
  <p:sink/>
  
  <p:add-attribute name="decorate-result-schematron" match="/*" 
    attribute-name="tr:rule-family" attribute-value="docx2hub">
    <p:input port="source">
      <p:pipe port="result-schematron" step="docx2hub"/>
    </p:input>
  </p:add-attribute>
  
  <p:sink/>
  
  <p:identity>
    <p:input port="source">
      <p:pipe port="result" step="pi"/>
    </p:input>
  </p:identity>
  
  <p:group use-when="doc-available('http://transpect.io/xproc-util/simple-progress-msg/xpl/simple-progress-msg.xpl')">
    <tr:simple-progress-msg name="success-msg" file="docx2hub-success.txt">
      <p:input port="msgs">
        <p:inline>
          <c:messages>
            <c:message xml:lang="en">Successfully finished DOCX to flat Hub XML conversion</c:message>
            <c:message xml:lang="de">Konvertierung von DOCX zu flachem Hub XML erfolgreich abgeschlossen</c:message>
          </c:messages>
        </p:inline>
      </p:input>
      <p:with-option name="status-dir-uri" select="$status-dir-uri"/>
    </tr:simple-progress-msg>
  </p:group>

</p:declare-step>
