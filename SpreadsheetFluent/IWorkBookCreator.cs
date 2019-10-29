using OfficeOpenXml;
using System.Collections.Generic;
namespace SpreadsheetFluent
{
    public interface IWorkBookCreator
    {
        byte[] Generate(params IEnumerable<object>[] datasets);
        ExcelPackage CreateExcel(params IEnumerable<object>[] datasets);
    }
}
