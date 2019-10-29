using OfficeOpenXml;
using SpreadsheetFluent.Blocks;
using SpreadsheetFluent.Enums;
using System;
using System.Collections.Generic;
namespace SpreadsheetFluent
{

    public abstract class AbstractWorksheet
    {
        internal IList<WorksheetBlockBase> Blocks { get; } = new List<WorksheetBlockBase>();
        internal Action<ExcelWorksheet> ExcelWorksheetOptions { get;  private set; }
        internal string Name { get; private set; }
        public AbstractWorksheet()
        {
            Name = Guid.NewGuid()
                        .ToString()
                        .Replace("-", "");
        }
        public IWorksheetBlock<T> CreateBlock<T>()
        {
            var block = new WorksheetBlock<T>();
            Blocks.Add(block);
            return block;
        }

        

        public AbstractWorksheet WorksheetName(string value)
        {
            Name = value;
            return this;
        }

        public AbstractWorksheet Configure(Action<ExcelWorksheet> opt)
        {
            ExcelWorksheetOptions = opt;
            return this;
        }

    
    }
}
