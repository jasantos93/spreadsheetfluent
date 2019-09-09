using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFluent.Test.Models
{
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public float Salary { get; set; }
        public DateTimeOffset Birthday { get; set; }
    }
}
