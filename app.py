"""FXS Intelligence • FX Funding Checker (Ministry of Industry tariff priorities)."""
from __future__ import annotations

import csv
import html
import io
import json
import re
from collections import Counter
from pathlib import Path

import streamlit as st

from engine import (
    ALLOWED_BY_PRIORITY, PATHS, diff_versions, evaluate_codes, export_csv,
    find_code, load_csv, load_version_table, normalize_hs, normalize_text,
    priority_label, result_rows, search, summary,
)
from excel_io import read_xlsx, write_xlsx

ROOT = Path(__file__).resolve().parent
DATA = ROOT / 'data'

st.set_page_config(page_title='FXS | FX Funding Checker', page_icon='💠', layout='wide', initial_sidebar_state='collapsed')
st.markdown('<style>' + (ROOT / 'styles.css').read_text(encoding='utf-8') + '</style>', unsafe_allow_html=True)


def esc(x):
    return html.escape(str(x),quote=True)


def persian_number(n):
    return f'{n:,}'.translate(str.maketrans('0123456789','۰۱۲۳۴۵۶۷۸۹'))


@st.cache_data(show_spinner=False)
def defaults():
    return load_csv(DATA / 'tariffs.csv'), json.loads((DATA / 'metadata.json').read_text(encoding='utf-8'))


def frame(title, subtitle=None):
    st.markdown(f'<div class="section-title"><h3>{esc(title)}</h3>{f"<p>{esc(subtitle)}</p>" if subtitle else ""}</div>', unsafe_allow_html=True)


def card(label, number, note='', accent=''):
    st.markdown(f'<div class="metric-card {accent}"><span>{esc(label)}</span><strong>{esc(number)}</strong><small>{esc(note)}</small></div>',unsafe_allow_html=True)


def banner(title, body, typ='info'):
    st.markdown(f'<div class="notice {typ}"><strong>{esc(title)}</strong><div>{esc(body)}</div></div>',unsafe_allow_html=True)


def dataframe(records, limit=400):
    if records:
        st.dataframe(records[:limit], use_container_width=True, hide_index=True, height=min(490,max(110,38*min(limit,len(records))+44)))
        if len(records)>limit:
            st.caption(f'برای سرعت نمایش، {persian_number(limit)} ردیف اول در جدول نمایش داده شده؛ خروجی شامل هر {persian_number(len(records))} ردیف است.')
    else:
        st.info('رکوردی مطابق فیلترها پیدا نشد.')


def buttons(records, prefix, label='دانلود نتایج'):
    if not records:
        return
    a,b,_=st.columns([1,1,3])
    with a:
        st.download_button(f'↓ {label} Excel',data=write_xlsx(records,title='FXS Result'),
                           file_name=f'{prefix}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           use_container_width=True, key=prefix+'_xlsx')
    with b:
        st.download_button('↓ دریافت CSV', data=export_csv(records),file_name=f'{prefix}.csv',
                           mime='text/csv',use_container_width=True,key=prefix+'_csv')


def read_table(file) -> list[list]:
    data=file.getvalue()
    if file.name.lower().endswith('.xlsx'):
        return read_xlsx(data, max_rows=20000)
    try:
        text=data.decode('utf-8-sig')
    except UnicodeDecodeError:
        text=data.decode('cp1256')
    sample=text[:6000]
    try:
        dialect=csv.Sniffer().sniff(sample,delimiters=',;\t')
        delimiter=dialect.delimiter
    except csv.Error:
        delimiter='\t' if '\t' in sample else ';' if sample.count(';')>sample.count(',') else ','
    table=list(csv.reader(io.StringIO(text),delimiter=delimiter))
    if len(table)>20000:
        raise ValueError('تعداد ردیف فایل بیش از سقف ۲۰٬۰۰۰ است.')
    return table


def extract_bulk(table):
    if not table:
        return []
    head=[normalize_text(x) for x in table[0]]
    aliases=['کد تعرفه','ردیف تعرفه','کد hs','hs_code','hs code','hscode','tariff','tariff code','code']
    col=next((i for i,v in enumerate(head) if v in aliases),None)
    if col is None:
        # Unheaded list: if first cell resembles an HS code include the first row.
        start=0 if normalize_hs(table[0][0],pad_short=True) else 1
        col=0
    else:
        start=1
    values=[]
    for row in table[start:]:
        if len(row)>col and row[col] is not None and str(row[col]).strip():
            values.append(row[col])
    return values


base_rows, metadata = defaults()
if 'active_snapshot' not in st.session_state:
    st.session_state.active_snapshot = None
if 'active_label' not in st.session_state:
    st.session_state.active_label = 'فایل مبنای ارسالی'
rows=st.session_state.active_snapshot if st.session_state.active_snapshot is not None else base_rows
stats=summary(rows)

st.markdown('''
<div class="tophead">
  <div class="identity"><div class="logo-mark">FXS</div><div><div class="wordmark">FXS Intelligence</div><div class="tagline">Helios Ecosystem <span class="sep">/</span> Foreign Exchange Intelligence</div></div></div>
  <span class="product-state">MINISTRY OF INDUSTRY · TARIFF DATA</span>
</div>
<div class="hero"><div class="eyebrow">HELIOS · TRADE INTELLIGENCE</div>
<h1>استعلام محل‌های مجاز تأمین ارز</h1>
<p>شناخت اولویت ارزی هر کد تعرفه و بررسی مسیرهای تأمین ارز در فرآیند ثبت سفارش وزارت صمت</p>
</div>
''',unsafe_allow_html=True)

b1,b2,b3=st.columns([2,2,3])
with b1: st.markdown(f'<div class="data-tag"><span class="live-dot"></span> {esc(st.session_state.active_label)}</div>',unsafe_allow_html=True)
with b2: st.markdown(f'<div class="data-tag">پوشش: {persian_number(stats["total"])} کد تعرفه</div>',unsafe_allow_html=True)
with b3: st.markdown('<div class="data-tag muted">دامنه: ثبت سفارش با مجوزدهنده ارزی وزارت صمت</div>',unsafe_allow_html=True)

tabs=st.tabs(['نمای کلی','استعلام تعرفه','جست‌وجوی معکوس','استعلام گروهی','نسخه‌ها و تغییرات'])

with tabs[0]:
    st.markdown('<div class="tabspace"></div>',unsafe_allow_html=True)
    frame('تصویر کلان اولویت‌های ارزی','وضعیت کدهای تعرفه در نسخه داده فعال')
    m1,m2,m3,m4=st.columns(4)
    with m1: card('مجموع کدهای تعرفه',persian_number(stats['total']),'کد یکتا در فایل')
    with m2: card('اولویت ۱ و ۲',persian_number(stats['by_priority'][1]+stats['by_priority'][2]),'مسیر صادرات دیگران مجاز','mint')
    with m3: card('فاقد اولویت اعلام‌شده',persian_number(stats['by_priority'][0]),'به معنی ممنوعیت واردات نیست','coral')
    with m4: card('دفاتر تخصصی',persian_number(len(stats['by_office'])),'پوشش وزارت صمت')
    c1,c2=st.columns([1.02,1.15],gap='large')
    with c1:
        frame('توزیع اولویت ارزی')
        blocks=[]
        maximum=max(stats['by_priority'].values()) or 1
        for p in [0,1,2,3,4]:
            count=stats['by_priority'][p]
            label='فاقد اولویت' if p==0 else f'اولویت {p}'
            bg='danger-bar' if p==0 else 'teal-bar'
            blocks.append(f'<div class="bar-line"><div class="bar-meta"><span>{label}</span><b>{persian_number(count)} <small>({100*count/max(1,stats["total"]):.1f}%)</small></b></div><div class="bar-track"><i class="{bg}" style="width:{count/maximum*100:.2f}%"></i></div></div>')
        st.markdown('<div class="visual-panel">'+''.join(blocks)+'</div>',unsafe_allow_html=True)
    with c2:
        frame('پوشش روش‌های تأمین ارز')
        maximum=max(stats['by_path'].values()) or 1
        blocks=[]
        for k,v in stats['by_path'].items():
            blocks.append(f'<div class="bar-line"><div class="bar-meta"><span>{esc(k)}</span><b>{persian_number(v)}</b></div><div class="bar-track"><i class="blue-bar" style="width:{v/maximum*100:.2f}%"></i></div></div>')
        st.markdown('<div class="visual-panel">'+''.join(blocks)+'</div>',unsafe_allow_html=True)
    frame('اولویت‌های تجاری','پیشنهادهای مستقیم برای اتصال قوانین به بازار')
    a,b=st.columns(2,gap='large')
    with a:
        banner('Exporter–Importer Matching',f'{persian_number(stats["by_path"][PATHS[1]])} تعرفه مسیر «صادرات دیگران» را دارند. این فقط ظرفیت تعرفه‌ای است، نه تعداد خریدار یا معامله قطعی.','teal')
    with b:
        banner('اهمیت اولویت صفر','برای این تعرفه‌ها هیچ‌یک از پنج روش این اطلاعیه مجاز اعلام نشده است؛ سایر ضوابط واردات مستقل است.','amber')
    with st.expander('تفکیک تعرفه‌ها بر اساس دفتر تخصصی',expanded=False):
        office_data=[{'دفتر تخصصی':k,'تعداد کد تعرفه':v} for k,v in stats['by_office'].most_common()]
        dataframe(office_data,100)

with tabs[1]:
    st.markdown('<div class="tabspace"></div>',unsafe_allow_html=True)
    frame('استعلام یک کد تعرفه','نتیجه برای کد ۸ رقمی دقیق محاسبه می‌شود؛ کدهای مشابه به‌صورت خودکار جایگزین نمی‌شوند.')
    col,_=st.columns([1.8,1])
    with col:
        query=st.text_input('کد تعرفه (HS8)',value='74031100',max_chars=30,placeholder='مثلاً 74031100',key='single_query')
    code=normalize_hs(query)
    if not query.strip():
        st.info('کد تعرفه ۸ رقمی را وارد کن.')
    elif not code:
        banner('کد نامعتبر','کد تعرفه باید دقیقاً شامل ۸ رقم باشد. استفاده از ارقام فارسی نیز مجاز است.','amber')
    else:
        hit=find_code(rows,code)
        if hit is None:
            banner('کد در این نسخه یافت نشد','وجود نداشتن کد در فایل به معنی ممنوعیت واردات یا فقدان اولویت در سایر منابع نیست.','amber')
            suggestions=[r for r in rows if r['hs_code'].startswith(code[:6])][:12]
            if suggestions:
                st.caption('کدهای نزدیک از نظر پیشوند (صرفاً پیشنهادی):')
                dataframe(result_rows(suggestions),12)
        else:
            a,b,c=st.columns([1.1,1.5,1],gap='medium')
            with a: card('کد تعرفه',hit['hs_code'],'HS · 8 DIGITS')
            with b: card('اولویت ارزی',priority_label(hit['priority']),'براساس فایل وزارت صمت','coral' if hit['priority']==0 else 'mint')
            with c: card('مسیر مجاز',persian_number(hit['funding_count'])+' از ۵','روش‌های مندرج در اطلاعیه')
            st.markdown(f'<div class="detail-head"><span>شرح کالا</span><h3>{esc(hit["description"] or "شرح کالا در فایل خالی است")}</h3><small>دفتر تخصصی: {esc(hit["office"] or "نامشخص")}</small></div>',unsafe_allow_html=True)
            frame('ماتریس استعلام محل تأمین ارز')
            chips=[]
            for path in PATHS:
                permitted=path in hit['allowed_paths']
                chips.append(f'<div class="funding-row {"ok" if permitted else "no"}"><span class="funding-icon">{"✓" if permitted else "×"}</span><span>{esc(path)}</span><b>{"مجاز" if permitted else "مجاز نیست"}</b></div>')
            st.markdown('<div class="funding-grid">'+''.join(chips)+'</div>',unsafe_allow_html=True)
            if hit['priority']==0:
                banner('فاقد محل تأمین ارز مجاز اعلام‌شده','طبق اطلاعیه فعلی، هیچ‌کدام از پنج مسیر برای این کد اعلام نشده است. این حکم به‌تنهایی معادل ممنوعیت واردات نیست.','amber')
            elif hit['matchmaking_eligible']:
                banner('سیگنال مچینگ صادرکننده و واردکننده','تأمین ارز از محل صادرات دیگران برای این تعرفه در فهرست مجاز است. احراز شرایط معامله، ثبت سفارش، تخصیص و سایر مجوزها جداگانه ضروری است.','teal')
            result=evaluate_codes(rows,[code])
            buttons(result, f'FXS_HS_{code}', 'گزارش')

with tabs[2]:
    st.markdown('<div class="tabspace"></div>',unsafe_allow_html=True)
    frame('جست‌وجوی معکوس','از مسیر تأمین ارز یا گروه کالایی به کدهای تعرفه واجد شرایط برس.')
    a,b=st.columns([1,1.15])
    with a:
        offices=['همه']+sorted(stats['by_office'].keys())
        office=st.selectbox('دفتر تخصصی',offices,key='office_filter')
    with b:
        path=st.selectbox('محل تأمین ارز', ['همه']+list(PATHS),key='path_filter')
    a,b=st.columns([1,1.15])
    with a:
        priority_choices={0:'۰ · فاقد اولویت',1:'۱ · اولویت یک',2:'۲ · اولویت دو',3:'۳ · اولویت سه',4:'۴ · اولویت چهار'}
        selected=st.multiselect('اولویت‌های ارزی',list(priority_choices.keys()),default=[0,1,2,3,4],format_func=lambda x:priority_choices[x],key='priority_filter')
    with b:
        term=st.text_input('جست‌وجو در HS Code یا شرح کالا',placeholder='مثلاً 7403، مس، کابل، شیر خشک ...',key='term_filter')
    found=search(rows,term=term,office=office,priorities=selected,path=path)
    st.markdown(f'<div class="results-count">{persian_number(len(found))} <small>کد تعرفه مطابق فیلترها</small></div>',unsafe_allow_html=True)
    results=result_rows(found)
    dataframe(results,350)
    buttons(results,'FXS_Filtered_Tariffs')

with tabs[3]:
    st.markdown('<div class="tabspace"></div>',unsafe_allow_html=True)
    frame('استعلام گروهی','فایل شامل یک ستون کد تعرفه را بارگذاری کن یا چند کد را مستقیم وارد کن.')
    a,b=st.columns([1.15,1],gap='large')
    with a:
        up=st.file_uploader('بارگذاری Excel یا CSV',type=['xlsx','csv'],key='bulk_file',help='نام ستون: کد تعرفه، ردیف تعرفه یا hs_code؛ در فایل تک‌ستونه بدون عنوان هم پشتیبانی می‌شود.')
    with b:
        paste=st.text_area('یا درج مستقیم کدها (هر خط یک کد)',height=116,placeholder='74031100\n39021010\n04021020',key='bulk_paste')
    values=[]
    if up is not None:
        try:
            values.extend(extract_bulk(read_table(up)))
        except Exception as exc:
            st.error(f'خطا در خواندن فایل: {exc}')
    if paste.strip():
        values.extend([p.strip() for p in re.split(r'[,;\n\t]+',paste) if p.strip()])
    if values:
        if len(values)>10000:
            st.error('حداکثر ۱۰٬۰۰۰ استعلام در هر بار پشتیبانی می‌شود.')
        else:
            records=evaluate_codes(rows,values)
            counts=Counter(r['نتیجه'] for r in records)
            a,b,c=st.columns(3)
            with a: card('کل استعلام‌ها',persian_number(len(records)),'با حفظ ترتیب و تکرار')
            with b: card('با مسیر مجاز',persian_number(counts['قابل استعلام / مسیرهای مجاز مشخص']),'حداقل یک روش مجاز','mint')
            with c: card('یافت‌نشده یا نامعتبر',persian_number(counts['کد در فایل موجود نیست']+counts['فرمت کد نامعتبر']),'نتیجه نامشخص','coral')
            dataframe(records,350)
            buttons(records,'FXS_Bulk_Results')
    else:
        banner('آماده دریافت فهرست','فایل تک‌ستونه یا فهرست کدها را وارد کن؛ ترتیب و کدهای تکراری برای ردیابی حفظ می‌شوند.','info')

with tabs[4]:
    st.markdown('<div class="tabspace"></div>',unsafe_allow_html=True)
    frame('کنترل نسخه‌های اطلاعیه','مقایسه فهرست جدید با فایل مبنا و شناسایی تغییر اولویت، اضافه‌شدن یا حذف تعرفه‌ها.')
    x,y=st.columns(2)
    with x:
        card('تعداد ردیف‌های فایل مبنا',persian_number(len(base_rows)),'نسخه همراه پروژه')
    with y:
        card('تعداد ردیف‌های نسخه فعال',persian_number(len(rows)),'فعال‌سازی تغییرات صرفاً در نشست فعلی','mint')
    with st.expander('مشخصات فایل مبنا و منبع',expanded=True):
        st.markdown(f'**نام منبع:** {metadata["name"]}')
        st.markdown('**تاریخ انتشار و تاریخ اجرای رسمی:** تأیید نشده در فایل ارسالی')
        st.markdown(f'**تاریخ ورود به پروژه:** {metadata["import_date"]}')
        st.code(metadata['source_sha256'],language=None)
        st.download_button('↓ دریافت فایل پیوست اصلی',data=(DATA/'source_announcement.xlsx').read_bytes(),
                           file_name='MIMT_FX_Funding_Original.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='source_file')
    new_file=st.file_uploader('آپلود فایل جدید وزارت صمت برای مقایسه',type=['xlsx','csv'],key='version_upload',help='فایل باید چهار ستون کد تعرفه، شرح فارسی، دفتر تخصصی و اولویت ارزی داشته باشد.')
    if new_file is not None:
        try:
            incoming=load_version_table(read_table(new_file))
            changed=diff_versions(rows,incoming)
            ct=Counter('اضافه‌شده' if c['نوع تغییر']=='اضافه‌شده' else 'حذف‌شده' if c['نوع تغییر']=='حذف‌شده از فایل' else 'ویرایش' for c in changed)
            col1,col2,col3=st.columns(3)
            with col1: card('رکوردهای جدید',persian_number(ct['اضافه‌شده']),'تعرفه اضافه‌شده')
            with col2: card('رکوردهای حذف‌شده',persian_number(ct['حذف‌شده']),'در نسخه جدید نیست')
            with col3: card('رکوردهای ویرایش‌شده',persian_number(ct['ویرایش']),'تغییر اولویت، دفتر یا شرح','mint')
            if changed:
                dataframe(changed,300)
                buttons(changed,'FXS_Tariff_Version_Diff')
            else:
                st.success('در مقایسه با نسخه فعال، تغییری در محتوا پیدا نشد.')
            a,b=st.columns([1,2])
            with a:
                if st.button('فعال‌سازی موقت نسخه جدید',type='primary',use_container_width=True,key='activate'):
                    st.session_state.active_snapshot=incoming
                    st.session_state.active_label=f'نسخه بارگذاری‌شده: {new_file.name}'
                    st.rerun()
            with b:
                st.caption('فعال‌سازی فقط در همین نشست انجام می‌شود؛ برای انتشار رسمی باید فایل CSV بسته نرم‌افزاری جایگزین و نسخه جدید Deploy شود.')
        except Exception as exc:
            st.error(f'فایل نسخه جدید پذیرفته نشد: {exc}')
    if st.session_state.active_snapshot is not None:
        if st.button('بازگشت به فایل مبنا',key='reset'):
            st.session_state.active_snapshot=None
            st.session_state.active_label='فایل مبنای ارسالی'
            st.rerun()

st.markdown('''<div class="bottomline"><b>FXS Intelligence</b> <span>·</span> HELIOS ECOSYSTEM <span>·</span> نسخه آزمایشی 1.0
<div>این سامانه ابزار تحلیل داده و اطلاعیه است؛ نتیجه جایگزین استعلام رسمی سامانه جامع تجارت، مجوز ثبت سفارش یا تأیید تأمین و تخصیص ارز نیست. <a href="https://crs.ntsw.ir/" target="_blank" rel="noopener noreferrer">سامانه مقررات تجاری ↗</a></div></div>''',unsafe_allow_html=True)
