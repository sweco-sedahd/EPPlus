using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Edate : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2, eErrorType.Value);
            double dateSerial = ArgToDecimal(arguments, 0);
            System.DateTime date = System.DateTime.FromOADate(dateSerial);
            int nMonthsToAdd = ArgToInt(arguments, 1);
            System.DateTime resultDate = date.AddMonths(nMonthsToAdd);
            return CreateResult(resultDate.ToOADate(), DataType.Date);
        }
    }
}