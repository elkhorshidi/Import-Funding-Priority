"""Small, dependency-free XLSX import/export helpers (OOXML, no openpyxl/pandas)."""
from __future__ import annotations

from io import BytesIO
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED, BadZipFile
import re

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _get_text(element):
    return ''.join(t.text or '' for t in element.findall('.//{%s}t' % S)) if element is not None else ''


def read_xlsx(data: bytes, *, max_rows: int = 20000) -> list[list]:
    if len(data) > 12 * 1024 * 1024:
        raise ValueError('حجم فایل ورودی بیش از حد مجاز است (۱۲ مگابایت).')
    try:
        with ZipFile(BytesIO(data)) as z:
            if sum(i.file_size for i in z.infolist()) > 60*1024*1024:
                raise ValueError('حجم محتوای بازشده فایل بیش از حد مجاز است.')
            strings=[]
            if 'xl/sharedStrings.xml' in z.namelist():
                root=ET.fromstring(z.read('xl/sharedStrings.xml'))
                strings=[_get_text(si) for si in root.findall('{%s}si'%S)]
            # Locate first sheet through relationship, not only a fixed sheet1 path.
            root=ET.fromstring(z.read('xl/workbook.xml'))
            sheet=root.find('.//{%s}sheet'%S)
            if sheet is None:
                raise ValueError('هیچ شیتی در فایل پیدا نشد.')
            relid=sheet.get('{%s}id'%R)
            rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
            rel=next((x for x in rels if x.get('Id')==relid),None)
            if rel is None:
                raise ValueError('مسیر شیت پیدا نشد.')
            target=rel.get('Target','').lstrip('/')
            if target.startswith('xl/'):
                spath=target
            else:
                import posixpath
                spath=posixpath.normpath(posixpath.join('xl',target))
            if not spath.startswith('xl/worksheets/') or spath not in z.namelist():
                raise ValueError('شیت مورد انتظار پیدا نشد.')
            content=ET.fromstring(z.read(spath))
            sheetdata=content.find('{%s}sheetData'%S)
            if sheetdata is None:
                return []
            result=[]
            for row in sheetdata.findall('{%s}row'%S):
                if len(result)>=max_rows:
                    raise ValueError(f'بیش از {max_rows:,} ردیف پشتیبانی نمی‌شود.')
                vals=[]
                for cell in row.findall('{%s}c'%S):
                    addr=cell.get('r','A1')
                    letters=re.match(r'[A-Z]+',addr)
                    if not letters:
                        continue
                    idx=0
                    for c in letters.group(0):
                        idx=idx*26+ord(c)-64
                    idx-=1
                    if idx > 60:
                        continue
                    while len(vals)<=idx:
                        vals.append(None)
                    ctype=cell.get('t','n')
                    v=cell.find('{%s}v'%S)
                    if ctype=='inlineStr':
                        value=_get_text(cell.find('{%s}is'%S))
                    elif ctype=='s' and v is not None:
                        value=strings[int(v.text)]
                    elif v is not None and v.text is not None:
                        raw=v.text
                        if ctype=='n':
                            try:
                                value=int(raw) if '.' not in raw and 'E' not in raw.upper() else float(raw)
                            except ValueError:
                                value=raw
                        else:
                            value=raw
                    else:
                        value=None
                    vals[idx]=value
                result.append(vals)
            return result
    except (KeyError,BadZipFile,ET.ParseError,ValueError) as exc:
        if isinstance(exc, ValueError):
            raise
        raise ValueError('فرمت Excel معتبر نیست یا فایل آسیب دیده است.') from exc


def _col_name(i: int) -> str:
    result=''
    while i:
        i,rem=divmod(i-1,26)
        result=chr(65+rem)+result
    return result


def _xml_text(v: object) -> str:
    value=str(v if v is not None else '')
    value=''.join(c for c in value if ord(c)>=32 or c in '\n\r\t')
    return escape(value,{'"':'&quot;'})


def write_xlsx(records: list[dict], *, title: str='نتایج استعلام') -> bytes:
    """Create a real Excel file with RTL, filters, freeze header and styled columns."""
    headers=list(records[0].keys()) if records else ['نتیجه']
    header_vals = ''.join(f'<c r="{_col_name(i)}1" t="inlineStr" s="1"><is><t>{_xml_text(h)}</t></is></c>' for i,h in enumerate(headers,1))
    body=[]
    for ri,row in enumerate(records,2):
        cells=[]
        for ci,h in enumerate(headers,1):
            value=row.get(h,'')
            cref=f'{_col_name(ci)}{ri}'
            # HS IDs stay strings even if they contain only digits.
            if isinstance(value,(float,int)) and not isinstance(value,bool):
                cells.append(f'<c r="{cref}" s="2"><v>{value}</v></c>')
            else:
                cells.append(f'<c r="{cref}" s="2" t="inlineStr"><is><t>{_xml_text(value)}</t></is></c>')
        body.append(f'<row r="{ri}">'+''.join(cells)+'</row>')
    cols=''.join(f'<col min="{i}" max="{i}" width="{min(58, max(14, 44 if "شرح" in h else 24 if "دفتر" in h else 19))}" customWidth="1"/>' for i,h in enumerate(headers,1))
    sheetxml=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="{S}"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>{cols}</cols><sheetData><row r="1" ht="31" customHeight="1">{header_vals}</row>{''.join(body)}</sheetData><autoFilter ref="A1:{_col_name(len(headers))}{max(1,len(records)+1)}"/></worksheet>'''
    content=f'''<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'''
    rels='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    workbook=f'''<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="{_xml_text(title[:31])}" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    wrels='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    styles=f'''<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{S}"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF123347"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="right" vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
    out=BytesIO()
    with ZipFile(out,'w',compression=ZIP_DEFLATED) as z:
        for name,text in [('[Content_Types].xml',content),('_rels/.rels',rels),('xl/workbook.xml',workbook),('xl/_rels/workbook.xml.rels',wrels),('xl/styles.xml',styles),('xl/worksheets/sheet1.xml',sheetxml)]:
            z.writestr(name,text)
    return out.getvalue()
