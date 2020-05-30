using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelReportGenerator.Models
{
    public class WorksheetColumnDefinition
    {
        public string ColumnName { get; set; }
        public Type DataType { get; set; }
        public Queue<LambdaExpression> InnerCollection { get; set; }
        public LambdaExpression DataMapping { get; set; }

        public WorksheetColumnDefinition(string columnName, Type dataType, Queue<LambdaExpression> innerCollection, LambdaExpression dataMapping)
        {
            ColumnName = columnName;
            DataType = dataType;
            InnerCollection = innerCollection;
            DataMapping = dataMapping;
        }
    }
}
