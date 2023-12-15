using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Days360 : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            double numDate1 = ArgToDecimal(arguments, 0);
            double numDate2 = ArgToDecimal(arguments, 1);
            System.DateTime dt1 = System.DateTime.FromOADate(numDate1);
            System.DateTime dt2 = System.DateTime.FromOADate(numDate2);

            var calcType = Days360Calctype.Us;
            if (arguments.Count() > 2)
            {
                bool european = ArgToBool(arguments, 2);
                if (european) calcType = Days360Calctype.European;
            }

            int startYear = dt1.Year;
            int startMonth = dt1.Month;
            int startDay = dt1.Day;
            int endYear = dt2.Year;
            int endMonth = dt2.Month;
            int endDay = dt2.Day;

            if (calcType == Days360Calctype.European)
            {
                if (startDay == 31) startDay = 30;
                if (endDay == 31) endDay = 30;
            }
            else
            {
                var calendar = new GregorianCalendar();
                int nDaysInFeb = calendar.IsLeapYear(dt1.Year) ? 29 : 28;

                // If the investment is EOM and (Date1 is the last day of February) and (Date2 is the last day of February), then change D2 to 30.
                if (startMonth == 2 && startDay == nDaysInFeb && endMonth == 2 && endDay == nDaysInFeb)
                {
                    endDay = 30;
                }

                // If the investment is EOM and (Date1 is the last day of February), then change D1 to 30.
                if (startMonth == 2 && startDay == nDaysInFeb)
                {
                    startDay = 30;
                }

                // If D2 is 31 and D1 is 30 or 31, then change D2 to 30.
                if (endDay == 31 && startDay is 30 or 31)
                {
                    endDay = 30;
                }

                // If D1 is 31, then change D1 to 30.
                if (startDay == 31)
                {
                    startDay = 30;
                }
            }

            int result = endYear * 12 * 30 + endMonth * 30 + endDay - (startYear * 12 * 30 + startMonth * 30 + startDay);
            return CreateResult(result, DataType.Integer);
        }

        private int GetNumWholeMonths(System.DateTime dt1, System.DateTime dt2)
        {
            System.DateTime startDate = new System.DateTime(dt1.Year, dt1.Month, 1).AddMonths(1);
            var endDate = new System.DateTime(dt2.Year, dt2.Month, 1);
            return (endDate.Year - startDate.Year) * 12 + (endDate.Month - startDate.Month);
        }

        private enum Days360Calctype
        {
            European,
            Us
        }
    }
}