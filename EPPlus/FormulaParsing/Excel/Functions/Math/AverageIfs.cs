/* Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 *
 * Author							Change						Date
 *******************************************************************************
 * Mats Alm   		                Added		                2015-02-01
 *******************************************************************************/

using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class AverageIfs : MultipleRangeCriteriasFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            FunctionArgument[] functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 3);
            List<ExcelDoubleCellValue> sumRange = ArgsToDoubleEnumerable(true, new List<FunctionArgument> { functionArguments[0] }, context).ToList();
            var argRanges = new List<ExcelDataProvider.IRangeInfo>();
            var criterias = new List<string>();
            for (int ix = 1; ix < 31; ix += 2)
            {
                if (functionArguments.Length <= ix) break;
                ExcelDataProvider.IRangeInfo rangeInfo = functionArguments[ix].ValueAsRangeInfo;
                argRanges.Add(rangeInfo);
                string value = functionArguments[ix + 1].Value?.ToString();
                criterias.Add(value);
            }

            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0]);
            IList<int> enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (int ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                List<int> indexes = GetMatchIndexes(argRanges[ix], criterias[ix]);
                matchIndexes = enumerable.Intersect(indexes);
            }

            double result = matchIndexes.Average(index => sumRange[index]);

            return CreateResult(result, DataType.Decimal);
        }
    }
}