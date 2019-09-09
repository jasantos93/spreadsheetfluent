using OfficeOpenXml.Style;
using SpreadsheetFluent.Blocks;
using System;
using System.Linq.Expressions;

namespace SpreadsheetFluent.Rules
{
    public interface IWorksheetBlockPropertyRuleStyle<T, TProperty> : IWorksheetBlockColumn<T>
    {
        IWorksheetBlockPropertyRuleStyle<T, TProperty> ForCaption(Action<ExcelStyle> opt);
        IWorksheetBlockPropertyRuleStyle<T, TProperty> ForCell(Action<ExcelStyle> opt);
    }

    public class WorksheetBlockPropertyRuleStyle<T, TProperty> : IWorksheetBlockPropertyRuleStyle<T, TProperty>
    {
        private readonly WorksheetBlockPropertyRule<T, TProperty> _worksheetBlockPropertyRule;

        public WorksheetBlockPropertyRuleStyle(WorksheetBlockPropertyRule<T, TProperty> worksheetBlockPropertyRule)
        {
            _worksheetBlockPropertyRule = worksheetBlockPropertyRule;
        }

        public IWorksheetBlockPropertyRuleStyle<T, TProperty> ForCaption(Action<ExcelStyle> opt)
        {
            return StyleOptions(Enums.WorksheetSection.Caption, opt);
        }

        public IWorksheetBlockPropertyRuleStyle<T, TProperty> ForCell(Action<ExcelStyle> opt)
        {
            return StyleOptions(Enums.WorksheetSection.Cell, opt);
        }

        #region -IWorksheetBlockPropertyRule Member
        public IWorksheetBlockPropertyRule<T, TProperty1> WithColumn<TProperty1>(Expression<Func<T, TProperty1>> exp) where TProperty1 : struct
        {
            return _worksheetBlockPropertyRule.WithColumn(exp);
        }

        public IWorksheetBlockPropertyRule<T, string> WithColumn(Expression<Func<T, string>> exp)
        {
            return _worksheetBlockPropertyRule.WithColumn(exp);
        }

        public IWorksheetBlockPropertyRule<T, DateTime> WithColumn(Expression<Func<T, DateTime>> exp)
        {
            return _worksheetBlockPropertyRule.WithColumn(exp);
        }
        #endregion

        #region -Helper
        private IWorksheetBlockPropertyRuleStyle<T, TProperty> StyleOptions(Enums.WorksheetSection section, Action<ExcelStyle> opts)
        {
            if (!_worksheetBlockPropertyRule.Rule.Styles.ContainsKey(section))
            {
                _worksheetBlockPropertyRule.Rule.Styles.Add(section, opts);
                return this;
            }

            _worksheetBlockPropertyRule.Rule.Styles[section] = opts;

            return this;
        }
        #endregion
    }

}
