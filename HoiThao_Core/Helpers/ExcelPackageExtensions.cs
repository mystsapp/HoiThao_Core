﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace HoiThao_Core.Helpers
{
    public static class ExcelPackageExtensions
    {
        public static DataTable ToDataTable(this ExcelPackage package)
        {
            //ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            //DataTable table = new DataTable();
            //foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
            //{
            //    table.Columns.Add(firstRowCell.Text);
            //}
            //for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            //{
            //    var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
            //    var newRow = table.NewRow();
            //    foreach (var cell in row)
            //    {
            //        newRow[cell.Start.Column - 1] = cell.Text;
            //    }
            //    table.Rows.Add(newRow);
            //}
            //return table;

            //var oSheet = package.Workbook.Worksheets["My Worksheet Name"];

            DataTable dt = new DataTable();
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1$"];

            //check if the worksheet is completely empty
            if (worksheet.Dimension == null)
            {
                return dt;
            }

            //create a list to hold the column names
            List<string> columnNames = new List<string>();

            //needed to keep track of empty column headers
            int currentColumn = 1;

            //loop all columns in the sheet and add them to the datatable
            foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                string columnName = cell.Text.Trim();

                //check if the previous header was empty and add it if it was
                if (cell.Start.Column != currentColumn)
                {
                    columnNames.Add("Header_" + currentColumn);
                    dt.Columns.Add("Header_" + currentColumn);
                    currentColumn++;
                }

                //add the column name to the list to count the duplicates
                columnNames.Add(columnName);

                //count the duplicate column names and make them unique to avoid the exception
                //A column named 'Name' already belongs to this DataTable
                int occurrences = columnNames.Count(x => x.Equals(columnName));
                if (occurrences > 1)
                {
                    columnName = columnName + "_" + occurrences;
                }

                //add the column to the datatable
                dt.Columns.Add(columnName);

                currentColumn++;
            }

            //start adding the contents of the excel file to the datatable
            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                DataRow newRow = dt.NewRow();

                //loop all cells in the row
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }

                dt.Rows.Add(newRow);
            }

            return dt;
        }
    }
}