using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SpreadsheetFluent.Enums;
using SpreadsheetFluent.Internal;
using System;
using System.Collections.Generic;

namespace SpreadsheetFluent.Blocks
{
    public abstract class WorksheetBlockDrawingBase
    {
        public string Name { get; set; }
        public Action<ExcelDrawing> Options { get; set; }
    }

    public abstract class WorksheetBlockBase
    {
        internal string Id { get; set; } = Guid.NewGuid().ToString();
        internal Type BlockType { get; set; }
        internal AbstractWorksheet AbstractWorksheet { get; }
        internal Dictionary<WorksheetSection, Action<ExcelStyle>> Styles { get; } = new Dictionary<WorksheetSection, Action<ExcelStyle>>();
        internal Lazy<string> Title { get;  set; }
        internal int StartRow { get; set; } = 1;
        internal int StartColumn { get; set; } = 1;
        internal IList<PropertyRule> Rules { get; private set; } = new List<PropertyRule>();
        internal void AddRule(PropertyRule rule) => Rules.Add(rule);
        internal T Clone<T>() where T: WorksheetBlockBase => (T)MemberwiseClone();
    }
}
