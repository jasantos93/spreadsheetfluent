using System.Collections.Generic;
namespace SpreadsheetFluent
{
    public interface IWorkBookCreator
    {
        byte[] Generate(params IEnumerable<object>[] datasets);
    }
}
