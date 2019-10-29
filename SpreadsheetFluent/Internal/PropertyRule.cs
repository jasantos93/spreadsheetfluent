using OfficeOpenXml;
using OfficeOpenXml.Style;
using SpreadsheetFluent.Enums;
using SpreadsheetFluent.Extensions;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace SpreadsheetFluent.Internal
{
    internal class PropertyRule
    {
        public Type PropertyType { get; }
        public LambdaExpression Expression { get; }

        public Func<object, object> GetValue { get; }
        public Lazy<string> Caption { get; set; }
        public Action<ExcelRange> AutoFit { get; set; }
        public Action<ExcelRange> Options { get; set; }

        public Dictionary<WorksheetSection, Action<ExcelStyle>> Styles { get; set; } = new Dictionary<WorksheetSection, Action<ExcelStyle>>();

        public PropertyRule(Func<object, object> propertyFunc, LambdaExpression exp, Type type)
        {
            PropertyType = type;
            GetValue = propertyFunc;
            Expression = exp;
        }

        public static PropertyRule Create<T, TProperty>(Expression<Func<T, TProperty>> exp)
        {
            var compiled = exp.Compile();
            return new PropertyRule(compiled.ToGeneric(), exp, typeof(TProperty));
        }
    }
}
