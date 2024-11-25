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
 * Mats Alm Added		                2016-12-27
 *******************************************************************************/

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing
{
    public class EpplusExcelDataProvider : ExcelDataProvider
    {
        private readonly ExcelPackage _package;
        private ExcelWorksheet _currentWorksheet;
        private Dictionary<ulong, INameInfo> _names = new();
        private RangeAddressFactory _rangeAddressFactory;

        public EpplusExcelDataProvider(ExcelPackage package)
        {
            _package = package;

            _rangeAddressFactory = new RangeAddressFactory(this);
        }

        public override int ExcelMaxColumns => ExcelPackage.MaxColumns;

        public override int ExcelMaxRows => ExcelPackage.MaxRows;

        public override ExcelNamedRangeCollection GetWorksheetNames(string worksheet)
        {
            ExcelWorksheet ws = _package.Workbook.Worksheets[worksheet];
            if (ws != null)
            {
                return ws.Names;
            }

            return null;
        }

        public override ExcelNamedRangeCollection GetWorkbookNameValues()
        {
            return _package.Workbook.Names;
        }

        public override IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            SetCurrentWorksheet(worksheet);
            string wsName = string.IsNullOrEmpty(worksheet) ? _currentWorksheet.Name : worksheet;
            ExcelWorksheet ws = _package.Workbook.Worksheets[wsName];
            return new RangeInfo(ws, fromRow, fromCol, toRow, toCol);
        }

        public override IRangeInfo GetRange(string worksheet, int row, int column, string address)
        {
            var addr = new ExcelAddress(worksheet, address);
            if (addr.Table != null)
            {
                addr = ConvertToA1C1(addr);
            }

            //SetCurrentWorksheet(addr.WorkSheet); 
            string wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
            ExcelWorksheet ws = _package.Workbook.Worksheets[wsName];
            //return new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            return new RangeInfo(ws, addr);
        }

        public override IRangeInfo GetRange(string worksheet, string address)
        {
            var addr = new ExcelAddress(worksheet, address);
            if (addr.Table != null)
            {
                addr = ConvertToA1C1(addr);
            }

            //SetCurrentWorksheet(addr.WorkSheet); 
            string wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
            ExcelWorksheet ws = _package.Workbook.Worksheets[wsName];
            //return new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            return new RangeInfo(ws, addr);
        }

        private ExcelAddress ConvertToA1C1(ExcelAddress addr)
        {
            //Convert the Table-style Address to an A1C1 address
            addr.SetRCFromTable(_package, addr);
            var a = new ExcelAddress(addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            a._ws = addr._ws;
            return a;
        }

        public override INameInfo GetName(string worksheet, string name)
        {
            ExcelNamedRange nameItem;
            ExcelWorksheet ws;
            if (string.IsNullOrEmpty(worksheet))
            {
                if (_package._workbook.Names.ContainsKey(name))
                {
                    nameItem = _package._workbook.Names[name];
                }
                else
                {
                    return null;
                }

                ws = null;
            }
            else
            {
                ws = _package._workbook.Worksheets[worksheet];
                if (ws != null && ws.Names.ContainsKey(name))
                {
                    nameItem = ws.Names[name];
                }
                else if (_package._workbook.Names.ContainsKey(name))
                {
                    nameItem = _package._workbook.Names[name];
                }
                else
                {
                    return null;
                }
            }

            ulong id = ExcelCellBase.GetCellID(nameItem.LocalSheetId, nameItem.Index, 0);

            if (_names.ContainsKey(id))
            {
                return _names[id];
            }

            var ni = new NameInfo
            {
                Id = id,
                Name = name,
                Worksheet = nameItem.Worksheet == null ? nameItem._ws : nameItem.Worksheet.Name,
                Formula = nameItem.Formula
            };
            if (nameItem._fromRow > 0)
            {
                ni.Value = new RangeInfo(nameItem.Worksheet ?? ws, nameItem._fromRow, nameItem._fromCol, nameItem._toRow, nameItem._toCol);
            }
            else
            {
                ni.Value = nameItem.Value;
            }

            _names.Add(id, ni);
            return ni;
        }

        public override IEnumerable<object> GetRangeValues(string address)
        {
            SetCurrentWorksheet(ExcelAddressInfo.Parse(address));
            var addr = new ExcelAddress(address);
            string wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
            ExcelWorksheet ws = _package.Workbook.Worksheets[wsName];
            return (IEnumerable<object>)new CellsStoreEnumerator<ExcelCoreValue>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
        }


        public object GetValue(int row, int column)
        {
            return _currentWorksheet.GetValueInner(row, column);
        }

        public bool IsMerged(int row, int column)
        {
            //return _currentWorksheet._flags.GetFlagValue(row, column, CellFlags.Merged);
            return _currentWorksheet.MergedCells[row, column] != null;
        }

        public bool IsHidden(int row, int column)
        {
            return _currentWorksheet.Column(column).Hidden || _currentWorksheet.Column(column).Width == 0 ||
                   _currentWorksheet.Row(row).Hidden || _currentWorksheet.Row(column).Height == 0;
        }

        public override object GetCellValue(string sheetName, int row, int col)
        {
            SetCurrentWorksheet(sheetName);
            return _currentWorksheet.GetValueInner(row, col);
        }

        public override ExcelCellAddress GetDimensionEnd(string worksheet)
        {
            ExcelCellAddress address = null;
            try
            {
                address = _package.Workbook.Worksheets[worksheet].Dimension.End;
            }
            catch
            {
            }

            return address;
        }

        private void SetCurrentWorksheet(ExcelAddressInfo addressInfo)
        {
            if (addressInfo.WorksheetIsSpecified)
            {
                _currentWorksheet = _package.Workbook.Worksheets[addressInfo.Worksheet];
            }
            else if (_currentWorksheet == null)
            {
                _currentWorksheet = _package.Workbook.Worksheets.First();
            }
        }

        private void SetCurrentWorksheet(string worksheetName)
        {
            if (!string.IsNullOrEmpty(worksheetName))
            {
                _currentWorksheet = _package.Workbook.Worksheets[worksheetName];
            }
            else
            {
                _currentWorksheet = _package.Workbook.Worksheets.First();
            }
        }

        //public override void SetCellValue(string address, object value)
        //{
        //    var addressInfo = ExcelAddressInfo.Parse(address);
        //    var ra = _rangeAddressFactory.Create(address);
        //    SetCurrentWorksheet(addressInfo);
        //    //var valueInfo = (ICalcEngineValueInfo)_currentWorksheet;
        //    //valueInfo.SetFormulaValue(ra.FromRow + 1, ra.FromCol + 1, value);
        //    _currentWorksheet.Cells[ra.FromRow + 1, ra.FromCol + 1].Value = value;
        //}

        public override void Dispose()
        {
            _package.Dispose();
        }

        public override string GetRangeFormula(string worksheetName, int row, int column)
        {
            SetCurrentWorksheet(worksheetName);
            return _currentWorksheet.GetFormula(row, column);
        }

        public override object GetRangeValue(string worksheetName, int row, int column)
        {
            SetCurrentWorksheet(worksheetName);
            return _currentWorksheet.GetValue(row, column);
        }

        public override string GetFormat(object value, string format)
        {
            ExcelStyles styles = _package.Workbook.Styles;
            ExcelNumberFormatXml.ExcelFormatTranslator ft = null;
            foreach (ExcelNumberFormatXml f in styles.NumberFormats)
            {
                if (f.Format == format)
                {
                    ft = f.FormatTranslator;
                    break;
                }
            }

            if (ft == null)
            {
                ft = new ExcelNumberFormatXml.ExcelFormatTranslator(format, -1);
            }

            return ExcelRangeBase.FormatValue(value, ft, format, ft.NetFormat);
        }

        public override List<Token> GetRangeFormulaTokens(string worksheetName, int row, int column)
        {
            return _package.Workbook.Worksheets[worksheetName]._formulaTokens.GetValue(row, column);
        }

        public override bool IsRowHidden(string worksheetName, int row)
        {
            bool b = _package.Workbook.Worksheets[worksheetName].Row(row).Height == 0 ||
                     _package.Workbook.Worksheets[worksheetName].Row(row).Hidden;

            return b;
        }

        public override void Reset()
        {
            _names = new Dictionary<ulong, INameInfo>(); //Reset name cache.            
        }

        public class CellInfo : ICellInfo
        {
            readonly CellsStoreEnumerator<ExcelCoreValue> _values;
            readonly ExcelWorksheet _ws;

            internal CellInfo(ExcelWorksheet ws, CellsStoreEnumerator<ExcelCoreValue> values)
            {
                _ws = ws;
                _values = values;
            }

            public string Address => _values.CellAddress;

            public int Row => _values.Row;

            public int Column => _values.Column;

            public string Formula => _ws.GetFormula(_values.Row, _values.Column);

            public object Value => _values.Value._value;

            public double ValueDouble => ConvertUtil.GetValueDouble(_values.Value._value, true);

            public double ValueDoubleLogical => ConvertUtil.GetValueDouble(_values.Value._value);

            public bool IsHiddenRow
            {
                get
                {
                    if (_ws.GetValueInner(_values.Row, 0) is RowInternal row)
                    {
                        return row.Hidden || row.Height == 0;
                    }

                    return false;
                }
            }

            public bool IsExcelError => ExcelErrorValue.Values.IsErrorValue(_values.Value._value);

            public IList<Token> Tokens => _ws._formulaTokens.GetValue(_values.Row, _values.Column);
        }

        public class NameInfo : INameInfo
        {
            public ulong Id { get; set; }
            public string Worksheet { get; set; }
            public string Name { get; set; }
            public string Formula { get; set; }
            public IList<Token> Tokens { get; internal set; }
            public object Value { get; set; }
        }

        public class RangeInfo : IRangeInfo
        {
            int _cellCount;
            readonly int _fromRow;
            readonly int _toRow;
            readonly int _fromCol;
            readonly int _toCol;
            readonly CellsStoreEnumerator<ExcelCoreValue> _values;
            internal ExcelWorksheet _ws;

            public RangeInfo(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
            {
                _ws = ws;
                _fromRow = fromRow;
                _fromCol = fromCol;
                _toRow = toRow;
                _toCol = toCol;
                Address = new ExcelAddressBase(_fromRow, _fromCol, _toRow, _toCol);
                Address._ws = ws.Name;
                _values = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, _fromRow, _fromCol, _toRow, _toCol);
                Current = new CellInfo(_ws, _values);
            }

            public RangeInfo(ExcelWorksheet ws, ExcelAddressBase address)
            {
                _ws = ws;
                _fromRow = address._fromRow;
                _fromCol = address._fromCol;
                _toRow = address._toRow;
                _toCol = address._toCol;
                Address = address;
                Address._ws = ws.Name;
                _values = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, _fromRow, _fromCol, _toRow, _toCol);
                Current = new CellInfo(_ws, _values);
            }

            public int GetNCells()
            {
                return (_toRow - _fromRow + 1) * (_toCol - _fromCol + 1);
            }

            public bool IsEmpty
            {
                get
                {
                    if (_cellCount > 0)
                    {
                        return false;
                    }

                    if (_values.Next())
                    {
                        _values.Reset();
                        return false;
                    }

                    return true;
                }
            }

            public bool IsMulti
            {
                get
                {
                    if (_cellCount == 0)
                    {
                        if (_values.Next() && _values.Next())
                        {
                            _values.Reset();
                            return true;
                        }

                        _values.Reset();
                        return false;
                    }

                    if (_cellCount > 1)
                    {
                        return true;
                    }

                    return false;
                }
            }

            public ICellInfo Current { get; }

            public ExcelWorksheet Worksheet => _ws;

            public void Dispose()
            {
                //_values = null;
                //_ws = null;
                //_cell = null;
            }

            object IEnumerator.Current => this;

            public bool MoveNext()
            {
                _cellCount++;
                return _values.MoveNext();
            }

            public void Reset()
            {
                _values.Init();
            }

            public IEnumerator<ICellInfo> GetEnumerator()
            {
                Reset();
                return this;
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return this;
            }

            public ExcelAddressBase Address { get; }

            public object GetValue(int row, int col)
            {
                return _ws.GetValue(row, col);
            }

            public object GetOffset(int rowOffset, int colOffset)
            {
                if (_values.Row < _fromRow || _values.Column < _fromCol)
                {
                    return _ws.GetValue(_fromRow + rowOffset, _fromCol + colOffset);
                }

                return _ws.GetValue(_values.Row + rowOffset, _values.Column + colOffset);
            }


            public bool NextCell()
            {
                _cellCount++;
                return _values.MoveNext();
            }
        }

        //public override void SetToTableAddress(ExcelAddress address)
        //{
        //    address.SetRCFromTable(_package, address);
        //}
    }
}