using System;

namespace ExcelReportGenerator.Sample.Model
{
    public class FakeController
    {
        public IActionResult File(byte[] vs, string v1, string v2)
        {
            throw new NotImplementedException();
        }

        public interface IActionResult
        {
        }

        public class FromServicesAttribute : Attribute
        {
        }
    }
}
