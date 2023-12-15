using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Large : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            FunctionArgument args = arguments.ElementAt(0);
            int index = ArgToInt(arguments, 1) - 1;
            IEnumerable<ExcelDoubleCellValue> values = ArgsToDoubleEnumerable(new List<FunctionArgument> { args }, context);
            ThrowExcelErrorValueExceptionIf(() => index < 0 || index >= values.Count(), eErrorType.Num);
            ExcelDoubleCellValue result = values.OrderByDescending(x => x).ElementAt(index);
            return CreateResult(result, DataType.Decimal);
        }
    }
}