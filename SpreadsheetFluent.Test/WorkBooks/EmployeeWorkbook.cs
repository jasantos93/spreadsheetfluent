using SpreadsheetFluent.Test.Worksheets;
using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFluent.Test.WorkBooks
{
    public class EmployeeWorkbook : AbstractWorkbook
    {
        public EmployeeWorkbook()
        {
            AddWorksheet<PersonInfoWorksheet>();
        }
    }
}
