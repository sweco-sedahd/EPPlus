﻿using System.Collections.Generic;
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public abstract class DateParsingFunction : ExcelFunction
    {
        protected System.DateTime ParseDate(IEnumerable<FunctionArgument> arguments, object dateObj)
        {
            var date = System.DateTime.MinValue;
            if (dateObj is string)
            {
                date = System.DateTime.Parse(dateObj.ToString(), CultureInfo.InvariantCulture);
            }
            else
            {
                double d = ArgToDecimal(arguments, 0);
                date = System.DateTime.FromOADate(d);
            }

            return date;
        }
    }
}