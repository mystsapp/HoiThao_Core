using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace HoiThao_Core.Helpers
{
    public class ReadAsDataTable
    {
        public static DataTable ExcelToDataTable(string fileName)
        {
            DataTable dataTable = new DataTable();
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    dataTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (Row row in rows)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        dataRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }

                    dataTable.Rows.Add(dataRow);
                }

            }
            dataTable.Rows.RemoveAt(0);

            return dataTable;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }

        public static DataTable READExcel(string path)

        {

            //Instance reference for Excel Application

            Microsoft.Office.Interop.Excel.Application objXL = null;

            //Workbook refrence

            Microsoft.Office.Interop.Excel.Workbook objWB = null;

            DataSet ds = new DataSet();

            try

            {

                //Instancing Excel using COM services

                objXL = new Microsoft.Office.Interop.Excel.Application();

                //Adding WorkBook

                objWB = objXL.Workbooks.Open(path);



                foreach (Microsoft.Office.Interop.Excel.Worksheet objSHT in objWB.Worksheets)

                {

                    int rows = objSHT.UsedRange.Rows.Count;

                    int cols = objSHT.UsedRange.Columns.Count;

                    DataTable dt = new DataTable("abc");

                    int noofrow = 1;



                    //If 1st Row Contains unique Headers for datatable include this part else remove it

                    //Start

                    for (int c = 1; c <= cols; c++)

                    {

                        string colname = objSHT.Cells[1, c].ToString();

                        dt.Columns.Add(colname);

                        noofrow = 2;

                    }

                    //END



                    for (int r = noofrow; r <= rows; r++)

                    {

                        DataRow dr = dt.NewRow();

                        for (int c = 1; c <= cols; c++)

                        {

                            dr[c - 1] = objSHT.Cells[r, c].ToString();

                        }

                        dt.Rows.Add(dr);

                    }

                    ds.Tables.Add(dt);



                }



                //Closing workbook

                objWB.Close();

                //Closing excel application

                objXL.Quit();



            }

            catch (Exception ex)

            {
                throw ex;
                objWB.Saved = true;

                //Closing work book

                objWB.Close();

                //Closing excel application

                objXL.Quit();

                //Response.Write("Illegal permission");

            }

            return ds.Tables["abc"];

        }
    }
}
