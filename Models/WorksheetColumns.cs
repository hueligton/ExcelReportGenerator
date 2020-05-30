using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelReportGenerator.Models
{
    public class WorksheetColumns
    {
        public ICollection<WorksheetColumnDefinition> Columns { get; set; }
        private Queue<LambdaExpression> InnerCollection { get; set; }
        public WorksheetColumns()
        {
            Columns = new List<WorksheetColumnDefinition>();
            InnerCollection = new Queue<LambdaExpression>();
        }
        public void AddColumn(string columnName, Type columnType)
        {
            Columns.Add(new WorksheetColumnDefinition(columnName, columnType, InnerCollection, null));
            InnerCollection = new Queue<LambdaExpression>();
        }

        public void AddColumn<TObject, TPropertyType>(string columnName, Expression<Func<TObject, TPropertyType>> dataMapping)
        {
            Columns.Add(new WorksheetColumnDefinition(columnName, typeof(TPropertyType), InnerCollection, dataMapping));
            InnerCollection = new Queue<LambdaExpression>();
        }

        public WorksheetColumns JoinCollection<TObject, TPropertyType>(Expression<Func<TObject, IEnumerable<TPropertyType>>> collectionMapping)
        {
            InnerCollection.Enqueue(collectionMapping);
            return this;
        }
    }
}
