import sys
import csv
import io
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from engine import (ALLOWED_BY_PRIORITY, PATHS, load_csv, normalize_hs, find_code,
                    summary, evaluate_codes, search, diff_versions, load_version_table)
from excel_io import read_xlsx, write_xlsx

class FXFundingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rows=load_csv(Path(__file__).resolve().parents[1]/'data'/'tariffs.csv')

    def test_source_counts(self):
        s=summary(self.rows)
        self.assertEqual(s['total'],8151)
        self.assertEqual(s['by_priority'],{0:3592,1:66,2:1677,3:1066,4:1750})
        self.assertEqual(len(s['by_office']),11)
        self.assertEqual(s['by_path'][PATHS[1]],1743)

    def test_priority_matrix(self):
        self.assertEqual([len(ALLOWED_BY_PRIORITY[i]) for i in range(5)],[0,5,4,3,2])
        self.assertNotIn(PATHS[1],ALLOWED_BY_PRIORITY[3])
        self.assertIn(PATHS[1],ALLOWED_BY_PRIORITY[2])

    def test_known_codes(self):
        self.assertEqual(find_code(self.rows,'۷۴۰۳۱۱۰۰')['priority'],2)
        self.assertEqual(find_code(self.rows,'04021020')['priority'],0)
        self.assertEqual(find_code(self.rows,'85044090')['priority'],3)
        self.assertIsNone(find_code(self.rows,'11111111'))

    def test_normalize_numeric_and_invalid(self):
        self.assertEqual(normalize_hs(4011000),'04011000')
        self.assertEqual(normalize_hs('۰۴۰۱۱۰۰۰'),'04011000')
        self.assertIsNone(normalize_hs('4011000'))
        self.assertEqual(normalize_hs('4011000',pad_short=True),'04011000')
        self.assertIsNone(normalize_hs('abcd'))

    def test_filter_path(self):
        m=search(self.rows,path=PATHS[1])
        self.assertEqual(len(m),1743)
        self.assertTrue(all(x['priority'] in (1,2) for x in m))

    def test_bulk(self):
        res=evaluate_codes(self.rows,['74031100','04021020','11111111','?','4011000'])
        self.assertEqual(len(res),5)
        self.assertEqual(res[0][PATHS[1]],'مجاز')
        self.assertEqual(res[1]['نتیجه'],'فاقد مسیر مجاز اعلام‌شده')
        self.assertEqual(res[2]['نتیجه'],'کد در فایل موجود نیست')
        self.assertEqual(res[3]['نتیجه'],'فرمت کد نامعتبر')
        self.assertEqual(res[4]['کد تعرفه'],'04011000')

    def test_diff(self):
        prior=[dict(self.rows[0]),dict(self.rows[1])]
        latest=[dict(prior[0])]
        latest[0]['priority']=4
        latest.append({'hs_code':'99999999','description':'تست','office':'تست','priority':1})
        diffs=diff_versions(prior,latest)
        self.assertEqual(len(diffs),3)
        self.assertEqual(sum('تغییر' in d['نوع تغییر'] for d in diffs),1)

    def test_official_import(self):
        p=Path(__file__).resolve().parents[1]/'data'/'source_announcement.xlsx'
        # Reader tests original XLSX fully; the first three are sufficient for parsing check.
        tab=read_xlsx(p.read_bytes())
        self.assertEqual(len(tab),8152)
        imported=load_version_table(tab)
        self.assertEqual(len(imported),8151)
        self.assertEqual(summary(imported)['by_priority'],summary(self.rows)['by_priority'])

    def test_export_excel_roundtrip_preserves_leading_zero(self):
        records=[{'کد تعرفه':'04021020','اولویت':0},{'کد تعرفه':'74031100','اولویت':2}]
        data=write_xlsx(records,title='Test')
        self.assertGreater(len(data),1000)
        cells=read_xlsx(data)
        self.assertEqual(cells[1][0],'04021020')
        self.assertEqual(cells[2][1],2)

if __name__=='__main__': unittest.main()
