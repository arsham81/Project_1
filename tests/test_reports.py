import tempfile
import unittest
from pathlib import Path
import openpyxl
from report_export import export_reports
from app import compare
from test_app import fixture

class ReportTests(unittest.TestCase):
 def test_excel_and_pdf_exports(self):
  a=fixture();b=fixture();result=compare(a,b)
  with tempfile.TemporaryDirectory() as d:
   x,p=export_reports(result,Path(d)/'report.xlsx',Path(d)/'report.pdf')
   book=openpyxl.load_workbook(x)
   self.assertEqual(len(book['Գրաֆիկներ']._images),2)
   self.assertEqual(book['Համեմատություն']['D5'].value,1.296)
   self.assertIn('Աղբյուրներ',book.sheetnames)
   self.assertEqual(book['Փաստ և պլան']['D2'].value,1)
   self.assertTrue(p.read_bytes().startswith(b'%PDF-'))
   self.assertGreater(p.stat().st_size,10000)

if __name__=='__main__':unittest.main()
