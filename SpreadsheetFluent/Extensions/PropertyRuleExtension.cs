using OfficeOpenXml;
using SpreadsheetFluent.Enums;
using SpreadsheetFluent.Internal;

namespace SpreadsheetFluent.Extensions
{
    internal static class PropertyRuleExtension
    {
        public static void ApplyStyle(this PropertyRule _, WorksheetSection section, ref ExcelRange range)
        {
            if (!_.Styles.ContainsKey(section)) return;
            _.Styles[section](range.Style);
        }
    }
}
