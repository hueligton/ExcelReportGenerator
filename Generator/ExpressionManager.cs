using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelReportGenerator.Generator
{
    public static class ExpressionManager
    {
        public static object AnaliseExpression<T>(T entity, Expression expression)
        {
            object returnValue = null;
            if (expression is MemberExpression)
            {
                returnValue = ExpressionManager.GetPropertyValue(entity, expression);
            }
            else if (expression is ConditionalExpression)
            {
                ConditionalExpression conditionalExpression = expression as ConditionalExpression;
                Expression test = conditionalExpression.Test;
                bool conditionalResult = false;
                if (test is MemberExpression)
                {
                    conditionalResult = (bool)ExpressionManager.GetPropertyValue(entity, test);
                }
                else if (test is BinaryExpression)
                {
                    BinaryExpression binaryExpression = test as BinaryExpression;
                    object leftValue = ExpressionManager.GetPropertyValue(entity, binaryExpression.Left);
                    object rightValue = ExpressionManager.GetPropertyValue(entity, binaryExpression.Right);
                    conditionalResult = ExpressionManager.AnaliseBinaryExpression(binaryExpression, leftValue, rightValue);
                }
                returnValue = conditionalResult ? ExpressionManager.GetPropertyValue(entity, conditionalExpression.IfTrue) : ExpressionManager.GetPropertyValue(entity, conditionalExpression.IfFalse);
            }

            return returnValue;
        }

        public static object GetPropertyValue<T>(T entity, Expression expression)
        {
            switch (expression.NodeType)
            {
                case ExpressionType.MemberAccess:
                    Stack<MemberExpression> memberExpressions = GetMemberExpressions(new Stack<MemberExpression>(), expression as MemberExpression);
                    return GetPropertyInfo(memberExpressions, entity);
                case ExpressionType.Constant:
                    return ((ConstantExpression)expression).Value;
                default:
                    return null;
            }
        }

        public static object GetPropertyValue(Expression expression)
        {
            return GetPropertyValue<object>(null, expression);
        }

        public static Stack<MemberExpression> GetMemberExpressions(Stack<MemberExpression> memberExpressions, MemberExpression member)
        {
            memberExpressions.Push(member);
            if (member.Expression.NodeType == ExpressionType.Parameter)
                return memberExpressions;
            else
                return GetMemberExpressions(memberExpressions, (MemberExpression)member.Expression);
        }

        public static object GetPropertyInfo(Stack<MemberExpression> memberExpressions, object entity)
        {
            if (entity == null)
                return null;
            if (memberExpressions.Count > 0)
            {
                MemberExpression memberExpression = memberExpressions.Pop();
                var prop = (PropertyInfo)memberExpression.Member;
                return GetPropertyInfo(memberExpressions, prop.GetValue(entity));
            }
            return entity;
        }


        public static bool AnaliseBinaryExpression(BinaryExpression binaryExpression, object leftValue, object rightValue)
        {
            switch (binaryExpression.NodeType)
            {
                case ExpressionType.Coalesce:
                    //TODO implementar
                    return false;
                case ExpressionType.Equal:
                    return leftValue.Equals(rightValue);
                case ExpressionType.NotEqual:
                    return !leftValue.Equals(rightValue);
                case ExpressionType.GreaterThan:
                    return decimal.Parse(leftValue.ToString()) > decimal.Parse(rightValue.ToString());
                case ExpressionType.GreaterThanOrEqual:
                    return decimal.Parse(leftValue.ToString()) >= decimal.Parse(rightValue.ToString());
                case ExpressionType.LessThan:
                    return decimal.Parse(leftValue.ToString()) < decimal.Parse(rightValue.ToString());
                case ExpressionType.LessThanOrEqual:
                    return decimal.Parse(leftValue.ToString()) <= decimal.Parse(rightValue.ToString());
                case ExpressionType.AndAlso:
                    //TODO recursão
                    return false;
                case ExpressionType.OrElse:
                    //TODO recursão
                    return false;
                default:
                    return false;
            }
        }
    }
}
