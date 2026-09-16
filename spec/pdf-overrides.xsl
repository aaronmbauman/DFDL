<?xml version="1.0" encoding="UTF-8"?>
<!--
  PDF style overrides, layered over the ISO stylesheet.  mn2pdf accepts a
  single override file, so everything we change lives here.

  Ragged right.  GFD.240 sets its body text ragged right, not justified.  The ISO stylesheet
  justifies it, which stretches the spaces around DFDL's long property names
  until they are hard to read, and is worst in the deeply nested lists of
  Appendix F and Appendix G where the measure is already narrow.

  mn2pdf merges this file over the base stylesheet, so each attribute set has
  to repeat the attributes it is not changing or they would be dropped.
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:attribute-set name="p-style">
    <xsl:attribute name="text-align">left</xsl:attribute>
    <xsl:attribute name="margin-bottom">8pt</xsl:attribute>
    <xsl:attribute name="line-height">1.13</xsl:attribute>
  </xsl:attribute-set>

  <xsl:attribute-set name="note-style">
    <xsl:attribute name="text-align">left</xsl:attribute>
    <xsl:attribute name="role">Note</xsl:attribute>
    <xsl:attribute name="font-size">10pt</xsl:attribute>
    <xsl:attribute name="margin-top">8pt</xsl:attribute>
    <xsl:attribute name="margin-bottom">12pt</xsl:attribute>
  </xsl:attribute-set>

  <xsl:attribute-set name="termnote-style">
    <xsl:attribute name="text-align">left</xsl:attribute>
    <xsl:attribute name="role">Note</xsl:attribute>
    <xsl:attribute name="font-size">10pt</xsl:attribute>
    <xsl:attribute name="margin-top">8pt</xsl:attribute>
    <xsl:attribute name="margin-bottom">8pt</xsl:attribute>
  </xsl:attribute-set>

  <xsl:attribute-set name="example-p-style">
    <xsl:attribute name="text-align">left</xsl:attribute>
    <xsl:attribute name="font-size">10pt</xsl:attribute>
    <xsl:attribute name="margin-top">8pt</xsl:attribute>
    <xsl:attribute name="margin-bottom">8pt</xsl:attribute>
  </xsl:attribute-set>

  <xsl:attribute-set name="termexample-style">
    <xsl:attribute name="text-align">left</xsl:attribute>
    <xsl:attribute name="font-size">10pt</xsl:attribute>
    <xsl:attribute name="margin-top">8pt</xsl:attribute>
    <xsl:attribute name="margin-bottom">8pt</xsl:attribute>
  </xsl:attribute-set>

  <!--
    Keeping a table whole.  Metanorma's keep-lines-together attribute reaches
    the semantic XML, but the ISO stylesheet never reads it on a table, so a
    table splits across pages whatever the source asks for.  The two separator
    suppression matrices in Section 14.2 must not split: their meaning is
    carried by merged regions that span rows, and a page break cuts them in
    half.  This makes the attribute do what it says on a table, and leaves
    every table without it free to break as before.
  -->
  <xsl:attribute-set name="table-style">
    <xsl:attribute name="keep-together.within-column">
      <xsl:choose>
        <xsl:when test="@keep-lines-together = 'true'">always</xsl:when>
        <xsl:otherwise>auto</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="table-omit-footer-at-break">true</xsl:attribute>
    <xsl:attribute name="table-layout">fixed</xsl:attribute>
    <xsl:attribute name="border"><xsl:value-of select="$table-border"/></xsl:attribute>
  </xsl:attribute-set>

  <!--
    Keeping a row span whole.  Table 43 is taller than a page, so it has to
    break somewhere, and the break was landing inside a row span: the
    continuation page opened with rows whose Symbol, Presentation and Meaning
    cells were blank, because the cell that carries them had been cut.

    A row that does not cover every column is one such continuation, so it is
    tied to the row above and the break falls between groups instead.  This is
    opt-in per table, with class="keep-row-groups", because the translation
    table in Appendix D has a 23-row span that would not fit on a page at all.
  -->
  <xsl:attribute-set name="table-body-row-style" use-attribute-sets="table-row-style">
    <xsl:attribute name="keep-with-previous.within-page">
      <xsl:choose>
        <xsl:when test="ancestor::*[local-name()='table'][1]/@class = 'keep-row-groups'
                    and (count(*[local-name()='td'][not(@colspan)]) + sum(*[local-name()='td']/@colspan))
                        &lt; count(ancestor::*[local-name()='table'][1]/*[local-name()='colgroup']/*[local-name()='col'])">always</xsl:when>
        <xsl:otherwise>auto</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:attribute-set>

  <!--
    Framing code examples.  GFD.240 draws a thin box round its code blocks and
    fills it a light grey: 846 of its 860 code paragraphs carry a half-point
    border on all four sides, and the published PDF renders the inside of the
    box as #F3F3F3 against a white page.  The ISO stylesheet sets neither, so
    the examples ran into the surrounding prose with nothing to mark them off.
  -->
  <xsl:attribute-set name="sourcecode-style">
    <xsl:attribute name="background-color">#F3F3F3</xsl:attribute>
    <xsl:attribute name="border">0.5pt solid black</xsl:attribute>
    <xsl:attribute name="padding">4pt</xsl:attribute>
    <xsl:attribute name="white-space">pre</xsl:attribute>
    <xsl:attribute name="wrap-option">wrap</xsl:attribute>
    <xsl:attribute name="role">Code</xsl:attribute>
    <xsl:attribute name="font-family">Courier New, <xsl:value-of select="$font_noto_sans_mono"/></xsl:attribute>
    <xsl:attribute name="margin-bottom">12pt</xsl:attribute>
  </xsl:attribute-set>

</xsl:stylesheet>
