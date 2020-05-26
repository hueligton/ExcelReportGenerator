using ExcelReportGenerator.Interfaces;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelReportGenerator.Sample.Model
{
    public class SampleEntity : IReportable
    {
        public int Id { get; set; }
        public string Name { get; set; }
        [Required]
        [ForeignKey("NestedSampleEntity")]
        public int NestedSampleEntityId { get; set; }
        public virtual NestedSampleEntity NestedSampleEntity { get; set; }
        public DateTime RegisterDate { get; set; }
    }
}
