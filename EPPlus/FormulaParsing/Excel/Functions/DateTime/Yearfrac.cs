using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Yearfrac : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            FunctionArgument[] functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            double date1Num = ArgToDecimal(functionArguments, 0);
            double date2Num = ArgToDecimal(functionArguments, 1);
            if (date1Num > date2Num) //Switch to make date1 the lowest date
            {
                double t = date1Num;
                date1Num = date2Num;
                date2Num = t;
                FunctionArgument fa = functionArguments[1];
                functionArguments[1] = functionArguments[0];
                functionArguments[0] = fa;
            }

            System.DateTime date1 = System.DateTime.FromOADate(date1Num);
            System.DateTime date2 = System.DateTime.FromOADate(date2Num);

            int basis = 0;
            if (functionArguments.Length > 2)
            {
                basis = ArgToInt(functionArguments, 2);
                ThrowExcelErrorValueExceptionIf(() => basis is < 0 or > 4, eErrorType.Num);
            }

            ExcelFunction func = context.Configuration.FunctionRepository.GetFunction("days360");
            var calendar = new GregorianCalendar();
            switch (basis)
            {
                case 0:
                    double d360Result = System.Math.Abs(func.Execute(functionArguments, context).ResultNumeric);
                    // reproducing excels behaviour
                    if (date1.Month == 2 && date2.Day == 31)
                    {
                        int daysInFeb = calendar.IsLeapYear(date1.Year) ? 29 : 28;
                        if (date1.Day == daysInFeb) d360Result++;
                    }

                    return CreateResult(d360Result / 360d, DataType.Decimal);
                case 1:
                    return CreateResult(System.Math.Abs((date2 - date1).TotalDays / CalculateAcutalYear(date1, date2)), DataType.Decimal);
                case 2:
                    return CreateResult(System.Math.Abs((date2 - date1).TotalDays / 360d), DataType.Decimal);
                case 3:
                    return CreateResult(System.Math.Abs((date2 - date1).TotalDays / 365d), DataType.Decimal);
                case 4:
                    List<FunctionArgument> args = functionArguments.ToList();
                    args.Add(new FunctionArgument(true));
                    double? result = System.Math.Abs(func.Execute(args, context).ResultNumeric / 360d);
                    return CreateResult(result.Value, DataType.Decimal);
                default:
                    return null;
            }
        }

        private double CalculateAcutalYear(System.DateTime dt1, System.DateTime dt2)
        {
            var calendar = new GregorianCalendar();
            double perYear = 0d;
            int nYears = dt2.Year - dt1.Year + 1;
            for (int y = dt1.Year; y <= dt2.Year; ++y)
            {
                perYear += calendar.IsLeapYear(y) ? 366 : 365;
            }

            if (new System.DateTime(dt1.Year + 1, dt1.Month, dt1.Day) >= dt2)
            {
                nYears = 1;
                perYear = 365;
                if (calendar.IsLeapYear(dt1.Year) && dt1.Month <= 2)
                    perYear = 366;
                else if (calendar.IsLeapYear(dt2.Year) && dt2.Month > 2)
                    perYear = 366;
                else if (dt2 is { Month: 2, Day: 29 })
                    perYear = 366;
            }

            return perYear / nYears;
        }
    }
}