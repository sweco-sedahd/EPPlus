﻿using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Degrees : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double angle = ArgToDecimal(arguments, 0);
            double result = angle * 180 / System.Math.PI;
            return CreateResult(result, DataType.Decimal);
        }
    }
}