using ExcelReportGenerator.Application;
using ExcelReportGenerator.Interfaces;
using ExcelReportGenerator.Sample.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelReportGenerator.Sample
{
    public class SampleApplication : ReportGenerator
    {
        private List<SampleEntity> _mockContext { get; set; }

        public SampleApplication()
        {
            _mockContext = new List<SampleEntity>();
        }

        public override ExcelPackage GenerateExcelReport(DateTime? initialDate, DateTime? finalDate)
        {
            var tabelaQuery = _mockContext.AsQueryable();
            ICollection<IReportable> reportables = FilterByDate(tabelaQuery, initialDate, finalDate);

            DataTable dataTable = CreateTable("SampleTable");

            Dictionary<string, Type> columns = GetColumns();
            AddColumns(dataTable, columns);

            AddDataToTable(dataTable, reportables);

            return GenerateExcelPackage(true);
        }

        private Dictionary<string, Type> GetColumns()
        {
            return new Dictionary<string, Type>
            {
                { "Id", typeof(int) },
                { "Register Date", typeof(DateTime) },
                { "Name", typeof(string) },
                { "City", typeof(string) },
                { "Estate", typeof(string) },
                { "Country", typeof(string) },
            };
        }

        protected override void AddDataToTable(DataTable table, ICollection<IReportable> reportables)
        {
            foreach (SampleEntity sampleEntity in reportables.OfType<SampleEntity>())
            {
                DataRow row = table.NewRow();
                row["Id"] = sampleEntity.Id;
                row["Register Date"] = sampleEntity.RegisterDate;
                row["Name"]=sampleEntity.Name;
                row["City"] = sampleEntity.NestedSampleEntity.City;
                row["Estate"] = sampleEntity.NestedSampleEntity.State;
                row["Country"] = sampleEntity.NestedSampleEntity.Country;
                table.Rows.Add(row);
            }
        }
    }
}
