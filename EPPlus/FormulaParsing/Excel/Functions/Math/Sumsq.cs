﻿using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Sumsq : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            double retVal = 0d;
            if (arguments != null)
            {
                foreach (FunctionArgument arg in arguments)
                {
                    retVal += Calculate(arg, context);
                }
            }

            return CreateResult(retVal, DataType.Decimal);
        }


        private double Calculate(FunctionArgument arg, ParsingContext context, bool isInArray = false)
        {
            double retVal = 0d;
            if (ShouldIgnore(arg))
            {
                return retVal;
            }

            if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (FunctionArgument item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    retVal += Calculate(item, context, true);
                }
            }
            else
            {
                if (arg.Value is ExcelDataProvider.IRangeInfo cs)
                {
                    foreach (ExcelDataProvider.ICellInfo c in cs)
                    {
                        if (ShouldIgnore(c, context) == false)
                        {
                            CheckForAndHandleExcelError(c);
                            retVal += System.Math.Pow(c.ValueDouble, 2);
                        }
                    }
                }
                else
                {
                    CheckForAndHandleExcelError(arg);
                    if (IsNumericString(arg.Value) && !isInArray)
                    {
                        double value = ConvertUtil.GetValueDouble(arg.Value);
                        return System.Math.Pow(value, 2);
                    }

                    bool ignoreBool = isInArray;
                    retVal += System.Math.Pow(ConvertUtil.GetValueDouble(arg.Value, ignoreBool), 2);
                }
            }

            return retVal;
        }
    }
}