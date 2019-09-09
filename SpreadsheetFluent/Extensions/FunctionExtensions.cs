using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFluent.Extensions
{
   internal static class FunctionExtensions
    {
        public static Func<object, object> ToGeneric<T, TProperty>(this Func<T, TProperty> func) => x => func((T)x);

    }
}
