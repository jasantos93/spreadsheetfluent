using OfficeOpenXml.Style;
using SpreadsheetFluent.Enums;
using SpreadsheetFluent.Rules;
using System;
using System.Linq.Expressions;

namespace SpreadsheetFluent.Blocks
{
    public interface IWorksheetBlockStyle<T> : IWorksheetBlockColumn<T>
    {
        IWorksheetBlockStyle<T> ForAll(Action<ExcelStyle> opt);
        IWorksheetBlockStyle<T> ForTitle(Action<ExcelStyle> opt);
        IWorksheetBlockStyle<T> ForHeader(Action<ExcelStyle> opt);
        IWorksheetBlockStyle<T> ForEachCell(Action<ExcelStyle> opt);
        IWorksheetBlockStyle<T> ForBody(Action<ExcelStyle> opt);
    }


    public class WorksheetBlockStyle<T> : IWorksheetBlockStyle<T>
    {
        private readonly WorksheetBlock<T> _worksheetBlock;
        public WorksheetBlockStyle(WorksheetBlock<T> worksheetBlock)
        {
            _worksheetBlock = worksheetBlock;
        }

        public IWorksheetBlockStyle<T> ForAll(Action<ExcelStyle> opt)
        {
            return StyleOptions(WorksheetSection.All, opt);
        }
        public IWorksheetBlockStyle<T> ForTitle(Action<ExcelStyle> opt)
        {
            return StyleOptions(WorksheetSection.Title, opt);
        }
        public IWorksheetBlockStyle<T> ForHeader(Action<ExcelStyle> opt)
        {
            return StyleOptions(WorksheetSection.Header, opt);
        }

        public IWorksheetBlockStyle<T> ForBody(Action<ExcelStyle> opt)
        {
            return StyleOptions(WorksheetSection.Body, opt);
        }

        public IWorksheetBlockStyle<T> ForEachCell(Action<ExcelStyle> opt)
        {
            return StyleOptions(WorksheetSection.EachCell, opt);
        }
        #region -IWorksheetBlockPropertyRule
        public IWorksheetBlockPropertyRule<T, TProperty> WithColumn<TProperty>(Expression<Func<T, TProperty>> exp) where TProperty : struct
        {
            return _worksheetBlock.WithColumn(exp);
        }

        public IWorksheetBlockPropertyRule<T, string> WithColumn(Expression<Func<T, string>> exp)
        {
            return _worksheetBlock.WithColumn(exp);
        }

        public IWorksheetBlockPropertyRule<T, DateTime> WithColumn(Expression<Func<T, DateTime>> exp)
        {
            return _worksheetBlock.WithColumn(exp);
        }
        #endregion

        #region -Helper
        private WorksheetBlockStyle<T> StyleOptions(WorksheetSection section, Action<ExcelStyle> opts)
        {
            if (!_worksheetBlock.Styles.ContainsKey(section))
            {
                _worksheetBlock.Styles.Add(section, opts);
                return this;
            }
            _worksheetBlock.Styles[section] = opts;
            return this;
        }


        #endregion
    }
}
