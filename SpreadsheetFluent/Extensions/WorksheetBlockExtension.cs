using OfficeOpenXml;
using SpreadsheetFluent.Blocks;
using SpreadsheetFluent.Enums;

namespace SpreadsheetFluent.Extensions
{
    internal static class WorksheetBlockExtension
    {
        public static void ApplyStyle(this WorksheetBlockBase _, WorksheetSection section, ref ExcelRange range)
        {
            if (!_.Styles.ContainsKey(section)) return;
            _.Styles[section](range.Style);
        }
    }
}
