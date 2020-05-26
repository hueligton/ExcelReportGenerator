using ExcelReportGenerator.Interfaces;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelReportGenerator.Application
{
    public abstract class ReportGenerator
    {
        private ICollection<DataTable> Tables { get; set; }

        protected ReportGenerator()
        {
            Tables = new List<DataTable>();
        }

        public abstract ExcelPackage GenerateExcelReport(DateTime? initialDate, DateTime? finalDate);

        protected abstract void AddDataToTable(DataTable table, ICollection<IReportable> reportables);

        protected DataTable CreateTable(string tableName)
        {
            var table = new DataTable(tableName);
            Tables.Add(table);
            return table;
        }

        protected ICollection<IReportable> FilterByDate(IQueryable<IReportable> queryableData, DateTime? initialDate, DateTime? finalDate)
        {
            ICollection<IReportable> entities;
            if (initialDate == null && finalDate == null)
            {
                entities = queryableData.ToList();
            }
            else if (initialDate == null)
            {
                entities = queryableData.Where(e => e.RegisterDate <= finalDate).ToList();
            }
            else if (finalDate == null)
            {
                entities = queryableData.Where(e => e.RegisterDate >= initialDate).ToList();
            }
            else
            {
                entities = queryableData.Where(e => e.RegisterDate >= initialDate && e.RegisterDate <= finalDate).ToList();
            }

            return entities;
        }

        protected void AddColumns(DataTable dataTable, Dictionary<string, Type> columns)
        {
            foreach (var column in columns)
            {
                dataTable.Columns.Add(column.Key, column.Value);
            }
        }

        protected ExcelPackage GenerateExcelPackage(bool saveFormatedDate)
        {
            var package = new ExcelPackage();
            foreach (var dataTable in Tables)
            {
                var worksheet = package.Workbook.Worksheets.Add(dataTable.TableName);
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, PrintHeaders: true);

                if (saveFormatedDate)
                    FormatTableData(dataTable, worksheet);
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
