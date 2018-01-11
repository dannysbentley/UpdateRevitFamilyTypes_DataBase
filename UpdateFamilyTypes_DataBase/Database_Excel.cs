using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/* To work with EPPlus library */
using OfficeOpenXml;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace UpdateFamilyTypes_DataBase
{
    class Database_Excel
    {
        public string FilePath { get; private set; }

        public void ExcelRequest(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentNullException("filePath");
            FilePath = filePath;
        }

        public List<ObjectWall> ExcelReadWallFile()
        {
            List<ObjectWall> ListOfWalls = new List<ObjectWall>();

            try
            {
                var package = new ExcelPackage(new FileInfo(FilePath));
                ExcelWorksheet workSheet = package.Workbook.Worksheets[1];

                var start = workSheet.Dimension.Start;
                var end = workSheet.Dimension.End;

                for (int row = 2; row <= end.Row; row++)
                {
                    ObjectWall Etabs_Object = new ObjectWall();
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        object cellValue = workSheet.Cells[row, col].Text;
                        CaseSwitch(col, cellValue, Etabs_Object);
                    }
                    ListOfWalls.Add(Etabs_Object);
                }

                // Close the Excel file. 
                package.Stream.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot open " + FilePath + " for reading. Exception raised - " + ex.Message);
            }
            return ListOfWalls;
        }


        public void CaseSwitch(int Column, object cellValue, ObjectWall wall)
        {
            try
            {
                switch (Column)
                {
                    case 1:
                        wall.Level = Convert.ToString(cellValue);
                        break;
                    case 2:
                        wall.SOMID = Convert.ToString(cellValue);
                        break;
                    case 3:
                        wall.Type = Convert.ToString(cellValue);
                        break;                    
                }
            }
            catch { }
        }
    }
}
