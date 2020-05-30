using ExcelReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelReportGenerator.Generator
{
    public static class ExpressionHelper
    {
        public static object GetValue<T>(T entity, WorksheetColumnDefinition columnDefinition)
        {
            Delegate compiledExpression = CompileExpression(typeof(T), columnDefinition.DataMapping.Body);
            return compiledExpression.DynamicInvoke(entity);
        }

        public static object GetValue<T>(T entity, Expression expression)
        {
            Delegate compiledExpression = CompileExpression(typeof(T), expression);
            return compiledExpression.DynamicInvoke(entity);
        }

        public static Delegate CompileExpression(Type entityType, Expression expression)
        {
            ParameterExpression pe = Expression.Parameter(entityType, "x");
            expression = new ParameterReplacer(pe).Visit(expression);
            return Expression.Lambda(expression, new ParameterExpression[] { pe }).Compile();
        }

        public static dynamic GetInnerCollection<T>(T entity, Queue<LambdaExpression> innerCollection)
        {
            if (innerCollection.Count > 0)
                return GetInnerCollection(GetValue(entity, innerCollection.Dequeue().Body), innerCollection);
            return entity;
        }
    }
}
