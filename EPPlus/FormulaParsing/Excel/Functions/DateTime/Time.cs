﻿/* Copyright (C) 2011  Jan Källman
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
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/

using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Time : TimeBaseFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            string firstArg = arguments.ElementAt(0).Value.ToString();
            if (arguments.Count() == 1 && TimeStringParser.CanParse(firstArg))
            {
                double result = TimeStringParser.Parse(firstArg);
                return new CompileResult(result, DataType.Time);
            }

            ValidateArguments(arguments, 3);
            int hour = ArgToInt(arguments, 0);
            int min = ArgToInt(arguments, 1);
            int sec = ArgToInt(arguments, 2);

            ThrowArgumentExceptionIf(() => sec is < 0 or > 59, "Invalid second: " + sec);
            ThrowArgumentExceptionIf(() => min is < 0 or > 59, "Invalid minute: " + min);
            ThrowArgumentExceptionIf(() => min < 0 || hour > 23, "Invalid hour: " + hour);


            double secondsOfThisTime = hour * 60 * 60 + min * 60 + sec;
            return CreateResult(GetTimeSerialNumber(secondsOfThisTime), DataType.Time);
        }
    }
}