using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;

namespace FireblanketDataProcessing
{
    public class SheetProcessor
    {
        public char[] ColumnCharacters = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

        public void writeSheetData(char column, int row, string data, ref IXLWorksheet sheet)
        {
            sheet.Cell(column + row.ToString()).Value = data;
        }

        public void writeSheetData(int column, int row, string data, ref IXLWorksheet sheet)
        {
            sheet.Cell(row, column).Value = data;
        }

        public void writeSheetData(char column, string[] data, ref IXLWorksheet sheet)
        {
            for (int i = 0; i < data.Count(); i++)
            {
                sheet.Cell(column + (i + 1).ToString()).Value = data[i];
            }
        }

        public void writeSheetData(char column, string[] data, int startFromRow, ref IXLWorksheet sheet)
        {
            for (int i = 0; i < data.Count(); i++)
            {
                sheet.Cell(column + (startFromRow + i + 1).ToString()).Value = data[i];
            }
        }

        public void makeBold(char column, int row, ref IXLWorksheet sheet)
        {
            sheet.Cell(column + row.ToString()).Style.Font.Bold = true;
        }

        public void makeUnderline(char column, int row, ref IXLWorksheet sheet)
        {
            sheet.Cell(column + row.ToString()).Style.Font.Underline = ClosedXML.Excel.XLFontUnderlineValues.Single;
        }

        public XLWorkbook createSheet(int numWorksheets, params string[] worksheetNames)
        {
            List<string> sheetNames = new List<string>();
            if (worksheetNames.Count() == 0)
            {
                for (int i = 0; i < numWorksheets; i++)
                {
                    sheetNames.Add(i.ToString());
                }
            }
            else
            {
                sheetNames.AddRange(worksheetNames);
            }

            var wb = new XLWorkbook();

            for (int i = 0; i < sheetNames.Count(); i++)
            {
                wb.Worksheets.Add(sheetNames[i]);
            }

            return wb;
        }

        public XLWorkbook createSheet(string sheetName)
        {

            var wb = new XLWorkbook();

            if (sheetName.Count() > 31)
            {
                sheetName = sheetName.Substring(0, 31);
            }

            wb.Worksheets.Add(sheetName);
            return wb;
        }

        public XLWorkbook addSheet(string sheetName, ref XLWorkbook book)
        {
            book.Worksheets.Add(sheetName);
            return book;
        }

        public bool saveSheet(string path, ref XLWorkbook book)
        {
            bool hasSucceded = false;

            while (!hasSucceded)
            {
                try
                {
                    book.SaveAs(path);
                    hasSucceded = true;
                }
                catch
                {
                    throw new Exception();
                }
            }

            return true;
        }

        public bool saveSheet(string path, bool makeNewFile, ref XLWorkbook book)
        {
            bool hasSucceded = false;

            while (!hasSucceded)
            {
                try
                {
                    using (File.Create(path)) { }
                    book.SaveAs(path);
                    hasSucceded = true;
                }
                catch
                {
                    throw new Exception();
                }
            }

            return true;
        }

        public void adjustColumnSize(ref IXLWorksheet sheet)
        {
            sheet.Columns().AdjustToContents();
        }

        //for some reason the characters for columns don't work correctly

        /*public string[] readSheetColumn(char column, ref IXLWorksheet sheet)
        {
            List<string> ls = new List<string>();

            for (int i = 0; i < sheet.Column(column).CellsUsed().Count(); i++)
            {
                string value = sheet.Column(column).Cell(i + 1).Value.ToString();
                ls.Add(value);
            }

            return ls.ToArray();
        }*/

        public string[] readSheetColumn(int column, ref IXLWorksheet sheet)
        {
            List<string> ls = new List<string>();

            for (int i = 0; i < sheet.Column(column).CellsUsed().Count(); i++)
            {
                string value = sheet.Column(column).Cell(i + 1).Value.ToString();
                ls.Add(value);
            }

            return ls.ToArray();
        }

        public string[] readSheetColumn(char column, int useSheet, ref XLWorkbook book)
        {
            IXLWorksheet sheet = book.Worksheet(useSheet);

            List<string> ls = new List<string>();

            for (int i = 0; i < sheet.Column(column).CellsUsed().Count(); i++)
            {
                string value = sheet.Column(column).Cell(i + 1).Value.ToString();
                ls.Add(value);
            }

            return ls.ToArray();
        }

        public string[] readSheetColumn(int column, int useSheet, ref XLWorkbook book)
        {
            IXLWorksheet sheet = book.Worksheet(useSheet);

            List<string> ls = new List<string>();

            for (int i = 0; i < sheet.Column(column).CellsUsed().Count(); i++)
            {
                string value = sheet.Column(column).Cell(i + 1).Value.ToString();
                ls.Add(value);
            }

            return ls.ToArray();
        }

        public string readSheetData(char column, int row, ref IXLWorksheet sheet)
        {
            string value = sheet.Column(column).Cell(row).Value.ToString();
            return value;
        }

        public string readSheetData(int column, int row, ref IXLWorksheet sheet)
        {
            string value = sheet.Column(column).Cell(row).Value.ToString();
            return value;
        }

        public string[] readSheetColumn(char column, int startFromRow, ref IXLWorksheet sheet)
        {
            List<string> sheetData = new List<string>();

            for (int i = 1; i <= sheet.Column(column).CellsUsed().Count(); i++)
            {
                string value = sheet.Column(column).Cell(i + (startFromRow - 1)).Value.ToString();
                sheetData.Add(value);
            }

            return sheetData.ToArray();
        }

        public string[] readSheetColumn(int column, int startFromRow, ref IXLWorksheet sheet)
        {
            List<string> sheetData = new List<string>();

            for (int i = 1; i <= sheet.Column(column).CellsUsed().Count(); i++)
            {
                string value = sheet.Column(column).Cell(i + (startFromRow - 1)).Value.ToString();
                sheetData.Add(value);
            }

            return sheetData.ToArray();
        }

        public string getSheetName(ref IXLWorksheet sheet)
        {
            string name = sheet.Name.ToString();
            return name;
        }

        public int getAmountCellsInColumn(char column, ref IXLWorksheet sheet)
        {
            int amount = sheet.Column(column).CellsUsed().Count();
            return amount;
        }

        public int getAmountCellsInColumn(int column, ref IXLWorksheet sheet)
        {
            int amount = sheet.Column(column).CellsUsed().Count();
            return amount;
        }

        public Tuple<char, int> findInstanceOfText(string text, ref IXLWorksheet sheet)
        {
            int amountOfColumns = sheet.ColumnCount();
            char column = '\0';
            int row = -1;

            for (int i = 0; i < amountOfColumns; i++)
            {
                for (int j = 0; j < getAmountCellsInColumn((char)i, ref sheet); j++)
                {
                    string value = readSheetData((char)i, j, ref sheet);
                    if (value == text)
                    {
                        column = ColumnCharacters[i + 1];
                        row = j;
                        break;
                    }
                }
            }

            if (column != '\0' && row != -1)
            {
                return new Tuple<char, int>(column, row);
            }
            else
            {
                return null;
            }
        }
    }
}