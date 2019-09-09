using System;
using System.Linq.Expressions;
using SpreadsheetFluent.Blocks;

namespace SpreadsheetFluent.Rules
{
    public abstract class WorksheetBlockRule<T, TProperty>: IWorksheetBlockColumn<T>
    {
        public WorksheetBlockRule(WorksheetBlock<T> worksheetBlock)
        {
            WorksheetBlock = worksheetBlock;
        }

        protected WorksheetBlock<T> WorksheetBlock { get; }

        public IWorksheetBlockPropertyRule<T, TProperty1> WithColumn<TProperty1>(Expression<Func<T, TProperty1>> exp) where TProperty1 : struct
        {
            return WorksheetBlock.WithColumn(exp);
        }

        public IWorksheetBlockPropertyRule<T, string> WithColumn(Expression<Func<T, string>> exp)
        {
            return WorksheetBlock.WithColumn(exp);
        }

        public IWorksheetBlockPropertyRule<T, DateTime> WithColumn(Expression<Func<T, DateTime>> exp)
        {
            return WorksheetBlock.WithColumn(exp);
        }
    }
}
