using System;
using System.Linq.Expressions;

namespace ExcelReportGenerator.Generator
{
    public static class ExpressionHelper
    {
        public static Delegate CompileExpression(Type entityType, Expression expression)
        {
            ParameterExpression pe = Expression.Parameter(entityType, "x");
            expression = new ParameterReplacer(pe).Visit(expression);
            return Expression.Lambda(expression, new ParameterExpression[] { pe }).Compile();
        }
    }
}
