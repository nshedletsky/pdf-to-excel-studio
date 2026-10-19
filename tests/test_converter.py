import io
import unittest
import fitz
import pandas as pd
from converter import extract_pdf, build_workbook, safe_sheet_name, infer_additive_columns

class ConverterTests(unittest.TestCase):
    def test_unique_sheet_names(self):
        used=set()
        self.assertNotEqual(safe_sheet_name('abc.pdf',used),safe_sheet_name('abc.pdf',used))
    def test_no_total_for_arbitrary_numbers(self):
        df=pd.DataFrame({'Identifier':[1,2], 'Value':[12,34]})
        self.assertEqual(infer_additive_columns(df,'Tables'),[])
    def test_program_credit_total(self):
        df=pd.DataFrame({'Course':['SPM 300','SPM 400'], 'Credits':[3,3]})
        out=build_workbook([('program.pdf',df,'Program sheet')])
        from zipfile import ZipFile
        with ZipFile(io.BytesIO(out)) as z:
            xml=z.read('xl/worksheets/sheet1.xml').decode()
            self.assertIn('SUM(',xml)
    def test_extract_selectable_pdf(self):
        doc=fitz.open()
        page=doc.new_page()
        page.insert_text((72,100),'SPM 300 Sports Event Management 3')
        page.insert_text((72,120),'SPM 400 Sports Facility Management 3')
        page.insert_text((72,140),'SPM 450 Sports Law 3')
        df,info=extract_pdf(doc.tobytes(),'sample.pdf')
        self.assertEqual(info['mode'],'Program sheet')
        self.assertEqual(len(df),3)
    def test_scanned_pdf_rejected(self):
        doc=fitz.open();doc.new_page()
        with self.assertRaises(ValueError):
            extract_pdf(doc.tobytes(),'empty.pdf')

if __name__=='__main__': unittest.main()
