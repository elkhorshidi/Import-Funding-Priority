"""Pure-Python rule engine for Helios FX Funding Checker.

Rules are intentionally scoped to Ministry of Industry, Mine and Trade (MIMT)
registration orders. No order authorization or allocation is inferred.
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

PATHS = (
    "خرید ارز از سیستم بانکی",
    "از محل صادرات دیگران",
    "از محل صادرات خود",
    "از محل ارز دیگران",
    "از محل ارز خود",
)

ALLOWED_BY_PRIORITY: dict[int, tuple[str, ...]] = {
    0: (),
    1: PATHS,
    2: PATHS[1:],
    3: PATHS[2:],
    4: PATHS[3:],
}

PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
REPLACEMENTS = str.maketrans({"ي": "ی", "ك": "ک", "ۀ": "ه", "ة": "ه", "ـ": " "})


def normalize_text(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).translate(PERSIAN_DIGITS).translate(REPLACEMENTS)
    text = re.sub(r"[\u200c\u200e\u200f\u202a-\u202e]", " ", text)
    return re.sub(r"\s+", " ", text).strip().casefold()


def normalize_hs(value: Any, *, pad_short: bool = False) -> str | None:
    """Validate eight-digit HS codes. Numeric spreadsheet cells may lose leading zero."""
    if value is None or isinstance(value, bool):
        return None
    from_numeric = isinstance(value, (int, float))
    if isinstance(value, float):
        if not value.is_integer():
            return None
        value = int(value)
    code = str(value).strip().translate(PERSIAN_DIGITS)
    code = code.replace("\u200c", "").replace(" ", "")
    if re.fullmatch(r"\d+\.0+", code):
        code = code.split(".")[0]
        from_numeric = True
    if not re.fullmatch(r"[0-9]+", code):
        return None
    if (from_numeric or pad_short) and len(code) < 8:
        code = code.zfill(8)
    return code if len(code) == 8 else None


def priority_label(priority: int | None) -> str:
    return "فاقد اولویت ارزی (۰)" if priority == 0 else f"اولویت {priority}" if priority in (1, 2, 3, 4) else "نامشخص"


def enrich(row: dict) -> dict:
    priority = int(row['priority'])
    result = dict(row)
    result['priority_label'] = priority_label(priority)
    result['allowed_paths'] = ALLOWED_BY_PRIORITY[priority]
    result['funding_count'] = len(result['allowed_paths'])
    result['matchmaking_eligible'] = "از محل صادرات دیگران" in result['allowed_paths']
    return result


def load_csv(content: bytes | str | Path) -> list[dict]:
    if isinstance(content, Path):
        content = content.read_bytes()
    if isinstance(content, bytes):
        content = content.decode('utf-8-sig')
    result: list[dict] = []
    seen: set[str] = set()
    reader = csv.DictReader(io.StringIO(content))
    if not set(('hs_code','description','office','priority')).issubset(set(reader.fieldnames or [])):
        raise ValueError('فایل CSV ساختار استاندارد hs_code, description, office, priority را ندارد.')
    for index, row in enumerate(reader, start=2):
        code = normalize_hs(row.get('hs_code'))
        if code is None:
            raise ValueError(f'کد تعرفه نامعتبر در ردیف {index}.')
        if code in seen:
            raise ValueError(f'کد تعرفه تکراری: {code}')
        seen.add(code)
        try:
            p = int(str(row.get('priority', '')).strip())
        except ValueError as exc:
            raise ValueError(f'اولویت نامعتبر برای {code}.') from exc
        if p not in ALLOWED_BY_PRIORITY:
            raise ValueError(f'اولویت خارج از محدوده (۰ تا ۴) برای {code}.')
        result.append({'hs_code': code, 'description': str(row.get('description') or '').strip(),
                       'office': str(row.get('office') or '').strip(), 'priority': p})
    return result


def load_version_table(table: list[list[Any]]) -> list[dict]:
    """Read official 4-column Persian XLSX / CSV into normalized snapshot."""
    if not table:
        raise ValueError('فایل خالی است.')
    header = [normalize_text(x) for x in table[0]]
    required = {
        'hs_code': ('ردیف تعرفه', 'کد تعرفه', 'hs_code', 'hs code', 'tariff code'),
        'description': ('شرح فارسی', 'شرح کالا', 'description'),
        'office': ('دفتر تخصصی', 'office'),
        'priority': ('اولویت ارزی وزارت صنعت، معدن و تجارت', 'اولویت ارزی', 'priority'),
    }
    index = {}
    for key, aliases in required.items():
        options = [normalize_text(z) for z in aliases]
        found = next((i for i, h in enumerate(header) if h in options), None)
        if found is None:
            raise ValueError(f'ستون ضروری «{aliases[0]}» در فایل پیدا نشد.')
        index[key] = found
    records=[]
    seen=set()
    for rownum, raw in enumerate(table[1:], 2):
        if not any(str(v or '').strip() for v in raw):
            continue
        def field(key):
            return raw[index[key]] if len(raw)>index[key] else None
        code=normalize_hs(field('hs_code'), pad_short=True)
        if code is None:
            raise ValueError(f'کد تعرفه نامعتبر در ردیف {rownum}.')
        if code in seen:
            raise ValueError(f'کد تعرفه تکراری در فایل جدید: {code}')
        seen.add(code)
        pv = field('priority')
        try:
            p = int(str(pv).strip())
        except (ValueError, TypeError) as exc:
            raise ValueError(f'اولویت نامعتبر در ردیف {rownum}؛ کد {code}.') from exc
        if p not in ALLOWED_BY_PRIORITY:
            raise ValueError(f'اولویت {p} برای کد {code} مجاز نیست (فقط ۰ تا ۴).')
        records.append({'hs_code':code, 'description':str(field('description') or '').strip(),
                        'office':str(field('office') or '').strip(), 'priority':p})
    if not records:
        raise ValueError('هیچ کد تعرفه معتبری یافت نشد.')
    return records


def summary(rows: list[dict]) -> dict:
    c=Counter(int(x['priority']) for x in rows)
    o=Counter(x['office'] or 'دفتر نامشخص' for x in rows)
    return {'total':len(rows),'by_priority':{i:c[i] for i in range(5)},
            'by_office':o,'by_path':{path:sum(c[p] for p,paths in ALLOWED_BY_PRIORITY.items() if path in paths) for path in PATHS}}


def find_code(rows: list[dict], value: Any) -> dict | None:
    code=normalize_hs(value)
    return next((enrich(r) for r in rows if r['hs_code']==code), None) if code else None


def search(rows: list[dict], *, term: str = '', office: str='همه', priorities: Iterable[int] = (0,1,2,3,4),
           path: str='همه') -> list[dict]:
    term=normalize_text(term)
    selected=set(int(p) for p in priorities)
    results=[]
    for row in rows:
        priority=int(row['priority'])
        if priority not in selected or (office!='همه' and row['office']!=office):
            continue
        if path!='همه' and path not in ALLOWED_BY_PRIORITY[priority]:
            continue
        if term and term not in normalize_text(row['hs_code']+' '+row['description']+' '+row['office']):
            continue
        results.append(row)
    return results


def evaluate_codes(rows: list[dict], values: Iterable[Any]) -> list[dict]:
    index={r['hs_code']:r for r in rows}
    results=[]
    for num, value in enumerate(values,1):
        code=normalize_hs(value,pad_short=True)
        if code is None:
            result={'ردیف':num,'کد ورودی':str(value),'کد تعرفه':'','شرح کالا':'', 'دفتر تخصصی':'',
                    'اولویت ارزی':'', 'نتیجه':'فرمت کد نامعتبر'}
            allowed=()
        elif code not in index:
            result={'ردیف':num,'کد ورودی':str(value),'کد تعرفه':code,'شرح کالا':'', 'دفتر تخصصی':'',
                    'اولویت ارزی':'', 'نتیجه':'کد در فایل موجود نیست'}
            allowed=()
        else:
            row=index[code]
            allowed=ALLOWED_BY_PRIORITY[row['priority']]
            result={'ردیف':num,'کد ورودی':str(value),'کد تعرفه':code, 'شرح کالا':row['description'],
                    'دفتر تخصصی':row['office'], 'اولویت ارزی':row['priority'],
                    'نتیجه':'فاقد مسیر مجاز اعلام‌شده' if row['priority']==0 else 'قابل استعلام / مسیرهای مجاز مشخص'}
        for path in PATHS:
            result[path]='مجاز' if path in allowed else 'غیرمجاز در این اطلاعیه' if result['نتیجه'] not in ('کد در فایل موجود نیست','فرمت کد نامعتبر') else 'نامشخص'
        results.append(result)
    return results


def diff_versions(old: list[dict], new: list[dict]) -> list[dict]:
    a={r['hs_code']:r for r in old}
    b={r['hs_code']:r for r in new}
    changes=[]
    for code in sorted(set(a)|set(b)):
        x,y=a.get(code),b.get(code)
        if x is None:
            typ='اضافه‌شده'
        elif y is None:
            typ='حذف‌شده از فایل'
        else:
            diffs=[]
            for col,label in (('priority','اولویت'),('office','دفتر تخصصی'),('description','شرح')):
                if x[col]!=y[col]:
                    diffs.append(label)
            if not diffs:
                continue
            typ='تغییر '+ '، '.join(diffs)
        changes.append({'کد تعرفه':code,'نوع تغییر':typ,
                        'اولویت قبلی':x['priority'] if x else '—',
                        'اولویت جدید':y['priority'] if y else '—',
                        'شرح قبلی':x['description'] if x else '—',
                        'شرح جدید':y['description'] if y else '—',
                        'دفتر قبلی':x['office'] if x else '—',
                        'دفتر جدید':y['office'] if y else '—'})
    return changes


def export_csv(records: list[dict]) -> bytes:
    if not records:
        return '\ufeff'.encode('utf-8')
    out=io.StringIO()
    writer=csv.DictWriter(out,fieldnames=list(records[0].keys()),extrasaction='ignore')
    writer.writeheader()
    writer.writerows(records)
    return ('\ufeff'+out.getvalue()).encode('utf-8')


def result_rows(rows: list[dict]) -> list[dict]:
    return [{'کد تعرفه':r['hs_code'],'شرح کالا':r['description'],'دفتر تخصصی':r['office'],
             'اولویت ارزی':r['priority'],**{path:'مجاز' if path in ALLOWED_BY_PRIORITY[r['priority']] else '—' for path in PATHS}}
            for r in rows]
