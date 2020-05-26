using System;

namespace ExcelReportGenerator.Sample
{
    class SampleController
    {
        public IActionResult GenerateExcelReport([FromServices] SampleApplication application, DateTime? initialDate, DateTime? finalDate)
        {
            using (var excel = application.GenerateExcelReport(initialDate, finalDate))
            {
                return File(excel.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ReportName.xlsx");
            }
        }

        #region ControllerBase and MVC Mock
        /*
         * This part of code should not be implemented.
         * This serve only as mock to the sample controller not fail.
         */

        private IActionResult File(byte[] vs, string v1, string v2)
        {
            throw new NotImplementedException();
        }

        public interface IActionResult
        {
        }

        private class FromServicesAttribute : Attribute
        {
        }
        #endregion
    }
}
