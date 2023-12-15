using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Networkdays : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            FunctionArgument[] functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            System.DateTime startDate = System.DateTime.FromOADate(ArgToInt(functionArguments, 0));
            System.DateTime endDate = System.DateTime.FromOADate(ArgToInt(functionArguments, 1));
            var calculator = new WorkdayCalculator();
            WorkdayCalculatorResult result = calculator.CalculateNumberOfWorkdays(startDate, endDate);
            if (functionArguments.Length > 2)
            {
                result = calculator.ReduceWorkdaysWithHolidays(result, functionArguments[2]);
            }

            return new CompileResult(result.NumberOfWorkdays, DataType.Integer);
        }
    }
}