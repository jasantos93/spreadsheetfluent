using OfficeOpenXml;
using SpreadsheetFluent.Blocks;
using SpreadsheetFluent.Enums;
using SpreadsheetFluent.Internal;
using System;
using System.Linq.Expressions;

namespace SpreadsheetFluent.Rules
{
    public interface IWorksheetBlockPropertyRule<T, TProperty> : IWorksheetBlockColumn<T>
    {
        IWorksheetBlockPropertyRule<T, TProperty> WithCaption(string value);
        IWorksheetBlockPropertyRule<T, TProperty> AutoFit();
        IWorksheetBlockPropertyRule<T, TProperty> AutoFit(double minimumWidth);
        IWorksheetBlockPropertyRule<T, TProperty> AutoFit(double minimumWidth, double maximumWidth);
        IWorksheetBlockPropertyRule<T, TProperty> Configure(Action<ExcelRange> options);
        IWorksheetBlockPropertyRuleStyle<T, TProperty> WithStyle();
    }

    public partial class WorksheetBlockPropertyRule<T, TProperty> : WorksheetBlockRule<T, TProperty>, IWorksheetBlockPropertyRule<T, TProperty>
    {
        internal PropertyRule Rule { get; }

        public WorksheetBlockPropertyRule(Expression<Func<T, TProperty>> expression, WorksheetBlock<T> worksheetBlock)
            : base(worksheetBlock)
        {
            Rule = PropertyRule.Create(expression);
            WorksheetBlock.AddRule(Rule);
        }
      
        public IWorksheetBlockPropertyRule<T, TProperty> WithCaption(string value)
        {
            Rule.Caption = new Lazy<string>(() => value);
            return this;
        }

        public IWorksheetBlockPropertyRule<T, TProperty> AutoFit()
        {
            Rule.AutoFit = (range) => range.AutoFitColumns();
            return this;
        }

        public IWorksheetBlockPropertyRule<T, TProperty> AutoFit(double minimumWidth)
        {
            Rule.AutoFit = (range) => range.AutoFitColumns(minimumWidth);
            return this;
        }
        public IWorksheetBlockPropertyRule<T, TProperty> AutoFit(double minimumWidth, double maximumWidth)
        {
            Rule.AutoFit = (range) => range.AutoFitColumns(minimumWidth, maximumWidth);
            return this;
        }

        public IWorksheetBlockPropertyRuleStyle<T, TProperty> WithStyle()
        {
            return new WorksheetBlockPropertyRuleStyle<T, TProperty>(this);
        }

        public IWorksheetBlockPropertyRule<T, TProperty> Configure(Action<ExcelRange> options)
        {
            Rule.Options = options;
            return this;
        }
    }
}
