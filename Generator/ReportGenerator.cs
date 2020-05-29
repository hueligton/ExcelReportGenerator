﻿using ExcelReportGenerator.Models;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Generator
{
    public class ReportGenerator
    {
        private Dictionary<string, WorksheetDefinition> WorkSheets { get; set; }

        public ReportGenerator()
        {
            WorkSheets = new Dictionary<string, WorksheetDefinition>();
        }


        public void CreateWorksheet(string worksheetName)
        {
            WorkSheets.Add(worksheetName, new WorksheetDefinition(worksheetName));
        }

        public void CreateWorksheet(string worksheetName, WorksheetColumns columns)
        {
            WorkSheets.Add(worksheetName, new WorksheetDefinition(worksheetName, columns));
        }

        public bool AddColumnsToWorksheet(string worksheetName, WorksheetColumns worksheetColumns)
        {
            if (WorkSheets.ContainsKey(worksheetName))
            {
                WorkSheets[worksheetName].WorksheetColumns = worksheetColumns;
                AddColumns(WorkSheets.GetValueOrDefault(worksheetName).Worksheet, worksheetColumns.Columns);
                return true;
            }
            return false;
        }

        private void AddColumns(DataTable dataTable, IList columns)
        {
            foreach (var column in columns)
            {
                WorksheetColumnDefinition columnDefinition = (WorksheetColumnDefinition)column;
                dataTable.Columns.Add(columnDefinition.ColumnName, columnDefinition.DataType);
            }
        }

        public bool AddDataToWorksheet<T>(string worksheetName, ICollection<T> reportables)
        {
            WorksheetDefinition worksheetDefinition;
            if (WorkSheets.ContainsKey(worksheetName))
                worksheetDefinition = WorkSheets.GetValueOrDefault(worksheetName);
            else
                return false;

            foreach (T entity in reportables)
            {
                DataRow row = worksheetDefinition.Worksheet.NewRow();
                foreach (var column in worksheetDefinition.WorksheetColumns.Columns)
                {
                    WorksheetColumnDefinition columnDefinition = (WorksheetColumnDefinition)column;
                    if (columnDefinition.DataMapping == null)
                        return false;

                    Delegate compiledExpression = ExpressionHelper.CompileExpression(typeof(T), columnDefinition.DataMapping.Body);
                    row[columnDefinition.ColumnName] = compiledExpression.DynamicInvoke(entity);
                }
                worksheetDefinition.Worksheet.Rows.Add(row);
            }
            return true;
        }

        public DataRow NewRow(string worksheetName)
        {
            if (WorkSheets.ContainsKey(worksheetName))
                return WorkSheets.GetValueOrDefault(worksheetName).Worksheet.NewRow();
            return null;
        }

        public void AddRow(string worksheetName, DataRow row)
        {
            if (WorkSheets.ContainsKey(worksheetName))
                WorkSheets.GetValueOrDefault(worksheetName).Worksheet.Rows.Add(row);
        }

        public ExcelPackage GenerateExcelPackage(bool saveFormattedData)
        {
            var package = new ExcelPackage();
            foreach (var worksheetDefinition in WorkSheets.Values)
            {
                var worksheet = package.Workbook.Worksheets.Add(worksheetDefinition.Worksheet.TableName);
                worksheet.Cells["A1"].LoadFromDataTable(worksheetDefinition.Worksheet, PrintHeaders: true);

                if (saveFormattedData)
                    FormatTableData(worksheetDefinition.Worksheet, worksheet);
            }
            return package;
        }

        private void FormatTableData(DataTable dataTable, ExcelWorksheet worksheet)
        {
            for (int linha = 0; linha < dataTable.Rows.Count; linha++)
            {
                for (int coluna = 0; coluna < dataTable.Columns.Count; coluna++)
                {
                    ConvertCellFormat(worksheet, linha + 2, coluna + 1, dataTable.Rows[linha], dataTable.Columns[coluna]);
                }
            }
        }

        private void ConvertCellFormat(ExcelWorksheet worksheet, int excelRow, int excelColumn, DataRow row, DataColumn column)
        {
            if (column.DataType == typeof(DateTime))
            {
                if (!row.IsNull(column))
                {
                    worksheet.SetValue(excelRow, excelColumn, (DateTime)row[column]);
                    var dtFormat = "dd/MM/yyyy";
                    worksheet.Cells[excelRow, excelColumn].Style.Numberformat.Format = dtFormat;
                }
                else
                    worksheet.SetValue(excelRow, excelColumn, null);
            }
            else
                worksheet.SetValue(excelRow, excelColumn, !row.IsNull(column) ? Convert.ChangeType(row[column].ToString(), column.DataType) : null);
        }
    }
}
