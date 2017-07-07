Excel file generator

The project is written in C# and the goal is to make a Console application that generates
an Excel file with name scores.xlsx. The sheet is populated with 100 rows of randomly generated data.
I have used Microsoft Excel 14.0 Object library.
I have created a class Excel for handling the operations with the Excel file (create, edit, save, close).
The directory for saving the file is determined by reflection and the file is saved in the directory of ExcelTest.exe file.
The Excel class has functions for reading cell, writing into cell, saving the file, saving file with name (saveAs), 
creating new file, creating new sheet and get a work sheed.
There are four columns: name, age, score and average score
For filling the names I have used a dictionary and randomizer so that evry 10 names are unique and are not repeating.
For the age and score columns I just used randomizer for randomizing the values.
For changing the colors for the odd rows I used For Loop. 

Code :


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

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelTest
{
    class Program
    {        
        static void Main()
        {
            //Create new file
            Excel ex = new Excel();
            ex.CreateNewFile();
            //Save file in the directory of the exe file
            ex.SaveAs(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\scores.xlsx");           

            //Write title columns            
            ex.WriteCell(1, 1, "Name");
            ex.WriteCell(1, 2, "Age");
            ex.WriteCell(1, 3, "Score");
            ex.WriteCell(1, 5, "Average Score");
            ex.WriteCell(2, 5, "=AVERAGE(C2:C101)");

            //Fill names
            FillNames(ex);

            //Fill age
            FillAge(ex);

            //Fill score
            FillScore(ex);

            //Set background colour and Bold the header
            Worksheet ws = ex.GetWorksheet;
            Range currentCell = ws.Cells[1, 1];
            currentCell.EntireRow.Font.Bold = true;
            currentCell.EntireRow.Interior.Color = System.Drawing.Color.LightSkyBlue;

            //Autofit text
            Range columns = ws.Columns;
            Range rows = ws.Rows;
            columns.AutoFit();
            rows.AutoFit();

            //Change text color for odd rows
            Range usedSpace = ws.UsedRange;
            Range usedRows = usedSpace.Rows;
            int counter = 1;
            foreach (Range row in usedRows)
            {
                if (counter > 1 && counter % 2 > 0)
                {
                    row.Characters.Font.Color = System.Drawing.Color.Green;
                }
                counter++;
            }

            //Save file
            ex.Save();

            //Close the excel application
            ex.Close();
            
        }
        
        public static void FillNames(Excel ex)
        {            
            List<string> names = new List<string>();
            int column = 1;
            string name;
            int currentRow = 1;
            for (int i = 0; i < 10; i++)
            {
                names = GenerateTenRandomNames();
                for (int j = 0; j < names.Count; j++)
                {
                    name = names[j];
                    currentRow = (i * 10) + j+2;
                    ex.WriteCell(currentRow, column, name);
                }
            }
        }

        public static void FillAge(Excel ex)
        {
            Random r = new Random();
            for (int i = 0; i < 100; i++)
            {
                ex.WriteCell((i + 2), 2, r.Next(20, 81).ToString());
            }
        }

        public static void FillScore(Excel ex)
        {
            Random r = new Random();
            for (int i = 0; i < 100; i++)
            {
                ex.WriteCell((i + 2), 3, r.Next(0, 101).ToString());
            }
        }
        //Function for generating 10 random ordered unique names from a list
        public static List<string> GenerateTenRandomNames()
        {
            Dictionary<int, string> dictionary = new Dictionary<int, string>();
            dictionary.Add(1, "Ivan");
            dictionary.Add(2, "Asen");
            dictionary.Add(3, "Petar");
            dictionary.Add(4, "Georgi");
            dictionary.Add(5, "Stoqn");
            dictionary.Add(6, "Nikolai");
            dictionary.Add(7, "Atanas");
            dictionary.Add(8, "Zdravko");
            dictionary.Add(9, "Bobi");
            dictionary.Add(10, "Damqn");

            Random r = new Random();
            HashSet<int> numbers = new HashSet<int>();
            List<string> names = new List<string>();
            string name;

            while (numbers.Count < 10)
            {
                int check = numbers.Count;
                int rInt = r.Next(1, 11);

                dictionary.TryGetValue(rInt, out name);
                numbers.Add(rInt);

                if (check < numbers.Count)
                {
                    names.Add(name);
                }
            }

            return names;
        }
    }
}


