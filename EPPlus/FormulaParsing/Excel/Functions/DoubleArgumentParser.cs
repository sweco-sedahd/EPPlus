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

using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;
using util = OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class DoubleArgumentParser : ArgumentParser
    {
        public override object Parse(object obj)
        {
            Require.That(obj).Named("argument").IsNotNull();
            if (obj is ExcelDataProvider.IRangeInfo)
            {
                ExcelDataProvider.ICellInfo r = ((ExcelDataProvider.IRangeInfo)obj).FirstOrDefault();
                return r?.ValueDouble ?? 0;
            }

            if (obj is double) return obj;
            if (obj.IsNumeric()) return util.ConvertUtil.GetValueDouble(obj);
            string str = obj != null ? obj.ToString() : string.Empty;
            try
            {
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                    return d;

                return System.DateTime.Parse(str, CultureInfo.CurrentCulture, DateTimeStyles.None).ToOADate();
            }
            catch // (Exception e)
            {
                throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
            }
        }
    }
}