﻿using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    /// <summary>
    /// Simple implementation of DateValue function, just using .NET built-in
    /// function System.DateTime.TryParse, based on current culture
    /// </summary>
    public class DateValue : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            string dateString = ArgToString(arguments, 0);
            return Execute(dateString);
        }

        internal CompileResult Execute(string dateString)
        {
            System.DateTime.TryParse(dateString, out System.DateTime result);
            return result != System.DateTime.MinValue ? CreateResult(result.ToOADate(), DataType.Date) : CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
        }
    }
}