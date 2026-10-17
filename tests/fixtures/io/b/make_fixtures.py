#!/usr/bin/env python3
"""Builds the hand-written fixtures in this folder from the published file format descriptions.

    python3 tests/fixtures/io/b/make_fixtures.py

Writes sample.sxc (OpenOffice.org 1.0 Calc), sample.ksp (KSpread, KoStore ZIP) and
sample-old.ksp (KSpread 1.x tar.gz store).
sample.sqlite is made the same way. All content is original test data. LibreOffice reads the .sxc and
.gnumeric fixtures, which is how they were checked (soffice --convert-to csv).
"""
import gzip, io, os, sqlite3, tarfile, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))


def write_zip(path, parts, mimetype=None):
    with zipfile.ZipFile(path, 'w') as z:
        if mimetype:
            z.writestr(zipfile.ZipInfo('mimetype'), mimetype, compress_type=zipfile.ZIP_STORED)
        for name, data in parts:
            z.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)


# ---------------------------------------------------------------------------------------------
# OpenOffice.org 1.0 Calc (.sxc)

OOO_NS = ('xmlns:office="http://openoffice.org/2000/office" xmlns:style="http://openoffice.org/2000/style" '
          'xmlns:text="http://openoffice.org/2000/text" xmlns:table="http://openoffice.org/2000/table" '
          'xmlns:draw="http://openoffice.org/2000/drawing" xmlns:fo="http://www.w3.org/1999/XSL/Format" '
          'xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" '
          'xmlns:meta="http://openoffice.org/2000/meta" xmlns:number="http://openoffice.org/2000/datastyle" '
          'xmlns:svg="http://www.w3.org/2000/svg" xmlns:config="http://openoffice.org/2001/config"')

SXC_CONTENT = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE office:document-content PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "office.dtd">
<office:document-content {OOO_NS} office:class="spreadsheet" office:version="1.0">
 <office:script/>
 <office:font-decls>
  <style:font-decl style:name="Arial" fo:font-family="Arial" style:font-family-generic="swiss" style:font-pitch="variable"/>
  <style:font-decl style:name="Courier New" fo:font-family="&apos;Courier New&apos;" style:font-family-generic="modern" style:font-pitch="fixed"/>
 </office:font-decls>
 <office:automatic-styles>
  <style:style style:name="co1" style:family="table-column"><style:properties fo:break-before="auto" style:column-width="2.267cm"/></style:style>
  <style:style style:name="co2" style:family="table-column"><style:properties fo:break-before="auto" style:column-width="4cm"/></style:style>
  <style:style style:name="ro1" style:family="table-row"><style:properties style:row-height="0.453cm" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>
  <style:style style:name="ro2" style:family="table-row"><style:properties style:row-height="0.8cm" fo:break-before="auto" style:use-optimal-row-height="false"/></style:style>
  <style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:properties table:display="true"/></style:style>
  <number:number-style style:name="N2" style:family="data-style"><number:number number:decimal-places="2" number:min-integer-digits="1"/></number:number-style>
  <number:date-style style:name="N37" style:family="data-style" number:automatic-order="true"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>
  <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:properties fo:background-color="#ffff00" fo:font-weight="bold" fo:text-align="center" style:text-align-source="fix" fo:border-bottom="0.002cm solid #000000"/></style:style>
  <style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N2"><style:properties fo:font-style="italic" fo:color="#0000ff"/></style:style>
  <style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>
  <style:style style:name="ce4" style:family="table-cell" style:parent-style-name="Default"><style:properties style:font-name="Courier New" fo:font-size="12pt" fo:wrap-option="wrap"/></style:style>
 </office:automatic-styles>
 <office:body>
  <table:calculation-settings table:null-year="1930"/>
  <table:table table:name="Fruit" table:style-name="ta1">
   <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
   <table:table-column table:style-name="co1" table:number-columns-repeated="4" table:default-cell-style-name="Default"/>
   <table:table-row table:style-name="ro2">
    <table:table-cell table:style-name="ce1"><text:p>Name</text:p><office:annotation office:author="Ann" office:create-date="2024-03-15"><text:p>Header note</text:p></office:annotation></table:table-cell>
    <table:table-cell table:style-name="ce1"><text:p>Qty</text:p></table:table-cell>
    <table:table-cell table:style-name="ce1"><text:p>Price</text:p></table:table-cell>
    <table:table-cell table:style-name="ce1"><text:p>Total</text:p></table:table-cell>
    <table:table-cell/>
   </table:table-row>
   <table:table-row table:style-name="ro1">
    <table:table-cell table:value-type="string"><text:p>Apple</text:p></table:table-cell>
    <table:table-cell table:value-type="float" table:value="3"><text:p>3</text:p></table:table-cell>
    <table:table-cell table:style-name="ce2" table:value-type="float" table:value="1.25"><text:p>1.25</text:p></table:table-cell>
    <table:table-cell table:style-name="ce2" table:formula="=[.B2]*[.C2]" table:value-type="float" table:value="3.75"><text:p>3.75</text:p></table:table-cell>
    <table:table-cell table:style-name="ce3" table:value-type="date" table:date-value="2024-03-15"><text:p>2024-03-15</text:p></table:table-cell>
   </table:table-row>
   <table:table-row table:style-name="ro1">
    <table:table-cell table:value-type="string"><text:p>Pear</text:p></table:table-cell>
    <table:table-cell table:value-type="float" table:value="5"><text:p>5</text:p></table:table-cell>
    <table:table-cell table:style-name="ce2" table:value-type="float" table:value="2.5"><text:p>2.50</text:p></table:table-cell>
    <table:table-cell table:style-name="ce2" table:formula="=[.B3]*[.C3]" table:value-type="float" table:value="12.5"><text:p>12.50</text:p></table:table-cell>
    <table:table-cell table:value-type="time" table:time-value="PT13H30M00S"><text:p>13:30:00</text:p></table:table-cell>
   </table:table-row>
   <table:table-row table:style-name="ro1">
    <table:table-cell table:value-type="string"><text:p>Plum</text:p></table:table-cell>
    <table:table-cell table:value-type="float" table:value="12"><text:p>12</text:p></table:table-cell>
    <table:table-cell table:style-name="ce2" table:value-type="float" table:value="0.75"><text:p>0.75</text:p></table:table-cell>
    <table:table-cell table:style-name="ce2" table:formula="=[.B4]*[.C4]" table:value-type="float" table:value="9"><text:p>9.00</text:p></table:table-cell>
    <table:table-cell table:value-type="boolean" table:boolean-value="true"><text:p>TRUE</text:p></table:table-cell>
   </table:table-row>
   <table:table-row table:style-name="ro1">
    <table:table-cell table:value-type="string"><text:p>Sum</text:p></table:table-cell>
    <table:table-cell table:number-columns-repeated="2"/>
    <table:table-cell table:style-name="ce2" table:formula="=SUM([.D2:.D4])" table:value-type="float" table:value="25.25"><text:p>25.25</text:p></table:table-cell>
    <table:table-cell table:style-name="ce4" table:value-type="string"><text:p>line one</text:p><text:p>line<text:s text:c="2"/>two</text:p></table:table-cell>
   </table:table-row>
   <table:table-row table:style-name="ro1">
    <table:table-cell table:number-columns-spanned="3" table:number-rows-spanned="1" table:value-type="string"><text:p>Merged title</text:p></table:table-cell>
    <table:covered-table-cell table:number-columns-repeated="2"/>
    <table:table-cell table:value-type="string"><text:p><text:a xlink:type="simple" xlink:href="https://example.com/">site</text:a></text:p></table:table-cell>
    <table:table-cell/>
   </table:table-row>
   <table:table-row table:style-name="ro1" table:visibility="collapse">
    <table:table-cell table:value-type="string"><text:p>hidden row</text:p></table:table-cell>
    <table:table-cell table:number-columns-repeated="4"/>
   </table:table-row>
  </table:table>
  <table:table table:name="Notes &amp; More" table:style-name="ta1">
   <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
   <table:table-row table:style-name="ro1"><table:table-cell table:value-type="string"><text:p>Tricky &lt;&amp;&gt; "quotes"</text:p></table:table-cell></table:table-row>
   <table:table-row table:style-name="ro1"><table:table-cell table:formula="=[$Fruit.D5]*2" table:value-type="float" table:value="50.5"><text:p>50.5</text:p></table:table-cell></table:table-row>
   <table:table-row table:style-name="ro1"><table:table-cell table:formula="=IF([$Fruit.B2]&gt;2;&quot;many&quot;;&quot;few&quot;)" table:value-type="string"><text:p>many</text:p></table:table-cell></table:table-row>
  </table:table>
  <table:named-expressions>
   <table:named-range table:name="Rate" table:base-cell-address="$Fruit.$A$1" table:cell-range-address="$Fruit.$C$2"/>
  </table:named-expressions>
 </office:body>
</office:document-content>
'''

SXC_STYLES = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles {OOO_NS} office:version="1.0">
 <office:font-decls><style:font-decl style:name="Arial" fo:font-family="Arial" style:font-family-generic="swiss" style:font-pitch="variable"/></office:font-decls>
 <office:styles>
  <style:default-style style:family="table-cell"><style:properties style:font-name="Arial" fo:font-size="10pt"/></style:default-style>
  <style:style style:name="Default" style:family="table-cell"><style:properties style:font-name="Arial"/></style:style>
 </office:styles>
</office:document-styles>
'''

SXC_META = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta {OOO_NS} office:version="1.0">
 <office:meta><meta:generator>OpenOffice.org 1.1.5 (Linux)</meta:generator><dc:title>Fruit list</dc:title><meta:initial-creator>Ann</meta:initial-creator><meta:creation-date>2004-03-15T10:00:00</meta:creation-date></office:meta>
</office:document-meta>
'''

SXC_SETTINGS = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings {OOO_NS} office:version="1.0">
 <office:settings>
  <config:config-item-set config:name="view-settings">
   <config:config-item-map-indexed config:name="Views">
    <config:config-item-map-entry>
     <config:config-item config:name="ActiveTable" config:type="string">Fruit</config:config-item>
     <config:config-item-map-named config:name="Tables">
      <config:config-item-map-entry config:name="Fruit">
       <config:config-item config:name="HorizontalSplitMode" config:type="short">0</config:config-item>
       <config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item>
       <config:config-item config:name="HorizontalSplitPosition" config:type="int">0</config:config-item>
       <config:config-item config:name="VerticalSplitPosition" config:type="int">1</config:config-item>
      </config:config-item-map-entry>
     </config:config-item-map-named>
    </config:config-item-map-entry>
   </config:config-item-map-indexed>
  </config:config-item-set>
 </office:settings>
</office:document-settings>
'''

SXC_MANIFEST = '''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE manifest:manifest PUBLIC "-//OpenOffice.org//DTD Manifest 1.0//EN" "Manifest.dtd">
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
 <manifest:file-entry manifest:media-type="application/vnd.sun.xml.calc" manifest:full-path="/"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>
</manifest:manifest>
'''


def make_sxc():
    write_zip(os.path.join(HERE, 'sample.sxc'), [
        ('content.xml', SXC_CONTENT), ('styles.xml', SXC_STYLES), ('meta.xml', SXC_META),
        ('settings.xml', SXC_SETTINGS), ('META-INF/manifest.xml', SXC_MANIFEST)], 'application/vnd.sun.xml.calc')


if __name__ == '__main__':
    make_sxc()
    print('fixtures written to', HERE)
