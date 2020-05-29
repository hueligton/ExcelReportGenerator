using ExcelReportGenerator.Sample.Model;
using System.Collections.Generic;

namespace ExcelReportGenerator.Sample
{
    class SampleController : FakeController
    {
        public IActionResult GenerateExcelReport([FromServices] SampleApplication application)
        {
            ICollection<SampleEntity> fakeSampleEntitiesCollection = new List<SampleEntity>();
            ICollection<AnotherSampleEntity> fakeAnotherSampleEntitiesCollection = new List<AnotherSampleEntity>();
            ICollection<OtherSampleEntity> fakeOtherSampleEntitiesCollection = new List<OtherSampleEntity>();

            using (var excel = application.GenerateExcelReport(fakeSampleEntitiesCollection, fakeAnotherSampleEntitiesCollection, fakeOtherSampleEntitiesCollection))
            {
                return File(excel.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ReportName.xlsx");
            }
        }
    }
}
