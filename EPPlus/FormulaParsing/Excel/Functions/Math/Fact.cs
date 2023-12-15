using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Fact : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double number = ArgToDecimal(arguments, 0);
            ThrowExcelErrorValueExceptionIf(() => number < 0, eErrorType.NA);
            double result = 1d;
            for (int x = 1; x < number; x++)
            {
                result *= x;
            }

            return CreateResult(result, DataType.Integer);
        }
    }
}