using System.Collections.Generic;

namespace ExcelReportGenerator.Sample.Model
{
    public class OtherSampleEntity
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public List<AnotherSampleEntity> AnotherSampleEntities { get; set; }
    }
}
