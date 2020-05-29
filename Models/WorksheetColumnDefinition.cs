using System;
using System.Linq.Expressions;

namespace ExcelReportGenerator.Models
{
    public class WorksheetColumnDefinition
    {
        public string ColumnName { get; set; }
        public Type DataType { get; set; }
        public LambdaExpression DataMapping { get; set; }

        public WorksheetColumnDefinition(string columnName, Type dataType, LambdaExpression dataMapping)
        {
            ColumnName = columnName;
            DataType = dataType;
            DataMapping = dataMapping;
        }
    }
}
