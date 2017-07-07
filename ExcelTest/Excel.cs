using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    class Excel
    {
        private string path = "";
        private _Application excel = new _Excel.Application();
        private Workbook wb;
        Worksheet ws;        

        public Excel()
        {
            
        }

        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];            
        }
        
        public Worksheet GetWorksheet
        {
            get
            {
                return this.ws;
            }           
        }        

        public string ReadCell (int i, int j)
        {
            if (ws.Cells[i,j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;
            }
            return "";
        }

        public void WriteCell (int i, int j, string s)
        {            
            ws.Cells[i, j].Value2 = s;
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            try
            {
                wb.SaveAs(path);
            }
            catch (COMException ex)
            {
                int error = ex.ErrorCode;                
            }
            
        }

        public void Close()
        {
            wb.Close(0);
            excel.Quit();
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }

        public void CreateNewSheet()
        {
            Worksheet tempSheet = wb.Worksheets.Add(After: ws);
        }
    }
}
