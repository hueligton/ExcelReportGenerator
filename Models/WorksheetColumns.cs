using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelReportGenerator.Models
{
    public class WorksheetColumns
    {
        public IList Columns { get; set; }
        public WorksheetColumns()
        {
            Columns = new List<object>();
        }

        public void AddColumn<TObject,TPropertyType>(string columnName, Expression<Func<TObject, TPropertyType>> dataMapping)
        {
            Columns.Add(new WorksheetColumnDefinition(columnName, typeof(TPropertyType), dataMapping));
        }

        public void AddColumn(string columnName, Type columnType)
        {
            Columns.Add(new WorksheetColumnDefinition(columnName, columnType, null));
        }
    }
}
