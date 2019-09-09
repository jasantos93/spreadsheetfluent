using System.Linq.Expressions;
using System.Reflection;

namespace SpreadsheetFluent.Extensions
{
    internal static class ExpressionExtensions
    {
        public static MemberInfo GetMember(this LambdaExpression expression)
        {
            if (!(RemoveUnary(expression.Body) is MemberExpression memberExp)) return null;
            var currentExpr = memberExp.Expression;
            while (true)
            {
                currentExpr = RemoveUnary(currentExpr);
                if (currentExpr != null && currentExpr.NodeType == ExpressionType.MemberAccess)
                    currentExpr = ((MemberExpression)currentExpr).Expression;
                else
                    break;
            }

            if (currentExpr == null || currentExpr.NodeType != ExpressionType.Parameter) return null;

            return memberExp.Member;
        }
        private static Expression RemoveUnary(Expression toUnwrap)
        {
            if (toUnwrap is UnaryExpression) return ((UnaryExpression)toUnwrap).Operand;
            return toUnwrap;
        }
    }
}
