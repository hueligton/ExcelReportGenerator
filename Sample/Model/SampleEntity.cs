using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelReportGenerator.Sample.Model
{
    public class SampleEntity
    {
        public int Id { get; set; }
        public string Name { get; set; }
        [Required]
        [ForeignKey("AnotherSampleEntity")]
        public int AnotherSampleEntityId { get; set; }
        public virtual AnotherSampleEntity AnotherSampleEntity { get; set; }
        public DateTime RegisterDate { get; set; }
    }
}
