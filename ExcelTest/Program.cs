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

            //Set background color and Bold the header
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
