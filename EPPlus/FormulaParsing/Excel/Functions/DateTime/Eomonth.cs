﻿using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Eomonth : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            System.DateTime date = System.DateTime.FromOADate(ArgToDecimal(arguments, 0));
            int monthsToAdd = ArgToInt(arguments, 1);
            System.DateTime resultDate = new System.DateTime(date.Year, date.Month, 1).AddMonths(monthsToAdd + 1).AddDays(-1);
            return CreateResult(resultDate.ToOADate(), DataType.Date);
        }
    }
}