using SpreadsheetFluent.Enums;
using SpreadsheetFluent.Rules;
using System;
using System.Linq.Expressions;

namespace SpreadsheetFluent.Blocks
{
    public interface IWorksheetBlockColumn<T>
    {
        IWorksheetBlockPropertyRule<T, TProperty> WithColumn<TProperty>(Expression<Func<T, TProperty>> exp) where TProperty : struct;
        IWorksheetBlockPropertyRule<T, string> WithColumn(Expression<Func<T, string>> exp);
        IWorksheetBlockPropertyRule<T, DateTime> WithColumn(Expression<Func<T, DateTime>> exp);
    }

    public interface IWorksheetBlock<T> : IWorksheetBlockColumn<T>
    {
        IWorksheetBlock<T> WithStartRow(int position);
        IWorksheetBlock<T> WithStartColumn(int position);
        IWorksheetBlock<T> WithTitle(string value);
        IWorksheetBlockStyle<T> WithStyle();
        IWorksheetBlock<T> WithDirection(BlockDirection direction);
    }

    public class WorksheetBlock<T> : WorksheetBlockBase, IWorksheetBlock<T>
    {
        public WorksheetBlock()
        {
            BlockType = typeof(T);
        }

        public IWorksheetBlockPropertyRule<T, TProperty> WithColumn<TProperty>(Expression<Func<T, TProperty>> expression) where TProperty : struct => WithColumnBase(expression);
        public IWorksheetBlockPropertyRule<T, string> WithColumn(Expression<Func<T, string>> exp) => WithColumnBase(exp);
        public IWorksheetBlockPropertyRule<T, DateTime> WithColumn(Expression<Func<T, DateTime>> exp)
        {
            return WithColumnBase(exp);
        }

        private IWorksheetBlockPropertyRule<T, TProperty> WithColumnBase<TProperty>(Expression<Func<T, TProperty>> exp)
        {
            return new WorksheetBlockPropertyRule<T, TProperty>(exp, this);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="position">The minimum value is 1</param>
        /// <returns></returns>
        public IWorksheetBlock<T> WithStartRow(int position)
        {
            StartRow = position + 1;
            return this;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="position">The minimum value is 1</param>
        /// <returns></returns>
        public IWorksheetBlock<T> WithStartColumn(int position)
        {
            StartColumn = position + 1;
            return this;
        }

        public IWorksheetBlockStyle<T> WithStyle()
        {
            return new WorksheetBlockStyle<T>(this);
        }
        public IWorksheetBlock<T> WithDirection(BlockDirection direction)
        {
            Direction = direction;
            return this;
        }
        public IWorksheetBlock<T> WithTitle(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return this;
            Title = new Lazy<string>(() => value);
            return this;
        }
    }
}
