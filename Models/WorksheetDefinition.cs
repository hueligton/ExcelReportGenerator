using System.Data;

namespace ExcelReportGenerator.Models
{
    public class WorksheetDefinition
    {
        public DataTable Worksheet { get; set; }
        public WorksheetColumns WorksheetColumns { get; set; }

        public WorksheetDefinition(string name)
        {
            Worksheet = new DataTable(name);
        }

        public WorksheetDefinition(string name, WorksheetColumns worksheetColumns) : this(name)
        {
            WorksheetColumns = worksheetColumns;
        }
    }
}
