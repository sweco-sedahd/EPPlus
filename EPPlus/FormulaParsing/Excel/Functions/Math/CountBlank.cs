﻿/* Copyright (C) 2011  Jan Källman
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
 * Author Change                      Date
 *******************************************************************************
 * Mats Alm Added		                2018-12-27
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class CountBlank : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            FunctionArgument arg = arguments.First();
            if (!arg.IsExcelRange) throw new InvalidOperationException("CountBlank only support ranges as arguments");
            int result = arg.ValueAsRangeInfo.GetNCells();
            foreach (ExcelDataProvider.ICellInfo cell in arg.ValueAsRangeInfo)
            {
                if (!(cell.Value == null || cell.Value.ToString() == string.Empty))
                {
                    result--;
                }
            }

            return CreateResult(result, DataType.Integer);
        }
    }
}