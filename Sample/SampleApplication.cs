using ExcelReportGenerator.Generator;
using ExcelReportGenerator.Models;
using ExcelReportGenerator.Sample.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelReportGenerator.Sample
{
    public class SampleApplication
    {
        public ExcelPackage GenerateExcelReport(ICollection<SampleEntity> sampleEntities, ICollection<AnotherSampleEntity> anotherSampleEntities, ICollection<OtherSampleEntity> otherSampleEntities)
        {
            ReportGenerator reportGenerator = new ReportGenerator();

            #region Worksheet sample 1
            reportGenerator.CreateWorksheet("SampleTable1");

            WorksheetColumns worksheetColumns1 = new WorksheetColumns();
            worksheetColumns1.AddColumn<SampleEntity, int>("Id", e => e.Id <= 5 || e.Id > 15 ? e.Id : 0);
            worksheetColumns1.AddColumn<SampleEntity, string>("Name", e => e.Name != null ? "DefaultName" : e.Name);
            worksheetColumns1.AddColumn<SampleEntity, DateTime>("Register Date", e => e.RegisterDate);
            worksheetColumns1.AddColumn<SampleEntity, string>("City", e => e.AnotherSampleEntity.City ?? "DefaultCity");
            worksheetColumns1.AddColumn<SampleEntity, string>("State", e => e.AnotherSampleEntity.State);
            worksheetColumns1.AddColumn<SampleEntity, string>("Country", e => e.AnotherSampleEntity.Country);
            worksheetColumns1.AddColumn<SampleEntity, string>("Active", e => e.AnotherSampleEntity.Active ? "Active" : "Inactive");

            reportGenerator.AddColumnsToWorksheet("SampleTable1", worksheetColumns1);
            reportGenerator.AddDataToWorksheet("SampleTable1", sampleEntities);
            #endregion Worksheet sample 1

            #region Worksheet sample 2
            WorksheetColumns worksheetColumns2 = new WorksheetColumns();
            worksheetColumns2.AddColumn<AnotherSampleEntity, string>("City", e => e.City);
            worksheetColumns2.AddColumn<AnotherSampleEntity, string>("State", e => e.State);
            worksheetColumns2.AddColumn<AnotherSampleEntity, string>("Country", e => e.Country);
            reportGenerator.CreateWorksheet("SampleTable2", worksheetColumns2);
            reportGenerator.AddDataToWorksheet("SampleTable2", anotherSampleEntities);
            #endregion Worksheet sample 2

            #region Worksheet sample 3
            reportGenerator.CreateWorksheet("Other Sample Table");

            WorksheetColumns worksheetColumns3 = new WorksheetColumns();
            worksheetColumns3.AddColumn("Entity Id", typeof(int));
            worksheetColumns3.AddColumn("Name", typeof(string));
            worksheetColumns3.AddColumn("City", typeof(string));
            worksheetColumns3.AddColumn("State", typeof(string));
            worksheetColumns3.AddColumn("Country", typeof(string));
            worksheetColumns3.AddColumn("Active", typeof(string));

            reportGenerator.AddColumnsToWorksheet("Other Sample Table", worksheetColumns3);
            AddDataToExcel(reportGenerator, "Other Sample Table", otherSampleEntities);
            #endregion Worksheet sample 3

            #region Worksheet sample 4
            reportGenerator.CreateWorksheet("Other Sample Table");
            WorksheetColumns worksheetColumns4 = new WorksheetColumns();
            worksheetColumns2.AddColumn<OtherSampleEntity, int>("Entity Id", e => e.Id);
            worksheetColumns2.AddColumn<OtherSampleEntity, string>("Name", e => e.Name);
            worksheetColumns4.JoinCollection<OtherSampleEntity, AnotherSampleEntity>(e => e.AnotherSampleEntities)
                .AddColumn<AnotherSampleEntity, string>("City", e => e.City);
            worksheetColumns4.JoinCollection<OtherSampleEntity, AnotherSampleEntity>(e => e.AnotherSampleEntities)
                .AddColumn<AnotherSampleEntity, string>("State", e => e.State);
            worksheetColumns4.JoinCollection<OtherSampleEntity, AnotherSampleEntity>(e => e.AnotherSampleEntities)
                .AddColumn<AnotherSampleEntity, string>("Country", e => e.Country);
            worksheetColumns4.JoinCollection<OtherSampleEntity, AnotherSampleEntity>(e => e.AnotherSampleEntities)
                .AddColumn<AnotherSampleEntity, string>("Active", e => e.Active ? "Active" : "Inactive");

            reportGenerator.AddColumnsToWorksheet("Other Sample Table", worksheetColumns3);
            AddDataToExcel(reportGenerator, "Other Sample Table", otherSampleEntities);
            #endregion Worksheet sample 4

            return reportGenerator.GenerateExcelPackage(saveFormattedData: true);
        }

        protected void AddDataToExcel(ReportGenerator reportGenerator, string worksheetName, ICollection<OtherSampleEntity> entities)
        {
            foreach (OtherSampleEntity otherSampleEntity in entities)
            {
                foreach (var anotherSampleEntity in otherSampleEntity.AnotherSampleEntities)
                {
                    DataRow row = reportGenerator.NewRow(worksheetName);
                    row["Entity Id"] = otherSampleEntity.Id;
                    row["Name"] = otherSampleEntity.Name;
                    row["City"] = anotherSampleEntity.City;
                    row["State"] = anotherSampleEntity.State;
                    row["Country"] = anotherSampleEntity.Country;
                    row["Active"] = anotherSampleEntity.Active ? "Active" : "Inactive";
                    reportGenerator.AddRow(worksheetName, row);
                }
            }
        }
    }
}