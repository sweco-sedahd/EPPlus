/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/

using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Bar chart
    /// </summary>
    public sealed class ExcelBarChart : ExcelChart
    {
        internal override eChartType GetChartType(string name)
        {
            if (name == "barChart")
            {
                if (Direction == eDirection.Bar)
                {
                    if (Grouping == eGrouping.Stacked)
                    {
                        return eChartType.BarStacked;
                    }

                    if (Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.BarStacked100;
                    }

                    return eChartType.BarClustered;
                }

                if (Grouping == eGrouping.Stacked)
                {
                    return eChartType.ColumnStacked;
                }

                if (Grouping == eGrouping.PercentStacked)
                {
                    return eChartType.ColumnStacked100;
                }

                return eChartType.ColumnClustered;
            }

            if (name == "bar3DChart")
            {
                #region "Bar Shape"

                if (Shape == eShape.Box)
                {
                    if (Direction == eDirection.Bar)
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.BarStacked3D;
                        }

                        if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.BarStacked1003D;
                        }

                        return eChartType.BarClustered3D;
                    }

                    if (Grouping == eGrouping.Stacked)
                    {
                        return eChartType.ColumnStacked3D;
                    }

                    if (Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.ColumnStacked1003D;
                    }

                    return eChartType.ColumnClustered3D;
                }

                #endregion

                #region "Cone Shape"

                if (Shape is eShape.Cone or eShape.ConeToMax)
                {
                    if (Direction == eDirection.Bar)
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.ConeBarStacked;
                        }

                        if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.ConeBarStacked100;
                        }

                        if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.ConeBarClustered;
                        }
                    }
                    else
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.ConeColStacked;
                        }

                        if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.ConeColStacked100;
                        }

                        if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.ConeColClustered;
                        }

                        return eChartType.ConeCol;
                    }
                }

                #endregion

                #region "Cylinder Shape"

                if (Shape == eShape.Cylinder)
                {
                    if (Direction == eDirection.Bar)
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.CylinderBarStacked;
                        }

                        if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.CylinderBarStacked100;
                        }

                        if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.CylinderBarClustered;
                        }
                    }
                    else
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.CylinderColStacked;
                        }

                        if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.CylinderColStacked100;
                        }

                        if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.CylinderColClustered;
                        }

                        return eChartType.CylinderCol;
                    }
                }

                #endregion

                #region "Pyramide Shape"

                if (Shape is eShape.Pyramid or eShape.PyramidToMax)
                {
                    if (Direction == eDirection.Bar)
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.PyramidBarStacked;
                        }

                        if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.PyramidBarStacked100;
                        }

                        if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.PyramidBarClustered;
                        }
                    }
                    else
                    {
                        if (Grouping == eGrouping.Stacked)
                        {
                            return eChartType.PyramidColStacked;
                        }

                        if (Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.PyramidColStacked100;
                        }

                        if (Grouping == eGrouping.Clustered)
                        {
                            return eChartType.PyramidColClustered;
                        }

                        return eChartType.PyramidCol;
                    }
                }

                #endregion
            }

            return base.GetChartType(name);
        }

        #region "Constructors"

        //internal ExcelBarChart(ExcelDrawings drawings, XmlNode node) :
        //    base(drawings, node/*, 1*/)
        //{
        //    SetChartNodeText("");
        //}
        //internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
        //    base(drawings, node, type)
        //{
        //    SetChartNodeText("");

        //    SetTypeProperties(drawings, type);
        //}
        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            SetChartNodeText("");

            SetTypeProperties(drawings, type);
        }

        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
            base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            SetChartNodeText(chartNode.Name);
        }

        internal ExcelBarChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
            SetChartNodeText(chartNode.Name);
        }

        #endregion

        #region "Private functions"

        //string _chartTopPath="c:chartSpace/c:chart/c:plotArea/{0}";
        private void SetChartNodeText(string chartNodeText)
        {
            if (string.IsNullOrEmpty(chartNodeText))
            {
                chartNodeText = GetChartNodeText();
            }
            //_chartTopPath = string.Format(_chartTopPath, chartNodeText);
            //_directionPath = string.Format(_directionPath, _chartTopPath);
            //_shapePath = string.Format(_shapePath, _chartTopPath);
        }

        private void SetTypeProperties(ExcelDrawings drawings, eChartType type)
        {
            /******* Bar direction *******/
            if (type is eChartType.BarClustered or eChartType.BarStacked or eChartType.BarStacked100 or eChartType.BarClustered3D or eChartType.BarStacked3D or eChartType.BarStacked1003D or eChartType.ConeBarClustered or eChartType.ConeBarStacked or eChartType.ConeBarStacked100 or eChartType.CylinderBarClustered or eChartType.CylinderBarStacked or eChartType.CylinderBarStacked100 or eChartType.PyramidBarClustered or eChartType.PyramidBarStacked or eChartType.PyramidBarStacked100)
            {
                Direction = eDirection.Bar;
            }
            else if (
                type is eChartType.ColumnClustered or eChartType.ColumnStacked or eChartType.ColumnStacked100 or eChartType.Column3D or eChartType.ColumnClustered3D or eChartType.ColumnStacked3D or eChartType.ColumnStacked1003D or eChartType.ConeCol or eChartType.ConeColClustered or eChartType.ConeColStacked or eChartType.ConeColStacked100 or eChartType.CylinderCol or eChartType.CylinderColClustered or eChartType.CylinderColStacked or eChartType.CylinderColStacked100 or eChartType.PyramidCol or eChartType.PyramidColClustered or eChartType.PyramidColStacked or eChartType.PyramidColStacked100)
            {
                Direction = eDirection.Column;
            }

            /****** Shape ******/
            if ( /*type == eChartType.ColumnClustered ||
                type == eChartType.ColumnStacked ||
                type == eChartType.ColumnStacked100 ||*/
                type == eChartType.Column3D ||
                type == eChartType.ColumnClustered3D ||
                type == eChartType.ColumnStacked3D ||
                type == eChartType.ColumnStacked1003D ||
                /*type == eChartType.BarClustered ||
                type == eChartType.BarStacked ||
                type == eChartType.BarStacked100 ||*/
                type == eChartType.BarClustered3D ||
                type == eChartType.BarStacked3D ||
                type == eChartType.BarStacked1003D)
            {
                Shape = eShape.Box;
            }
            else if (
                type is eChartType.CylinderBarClustered or eChartType.CylinderBarStacked or eChartType.CylinderBarStacked100 or eChartType.CylinderCol or eChartType.CylinderColClustered or eChartType.CylinderColStacked or eChartType.CylinderColStacked100)
            {
                Shape = eShape.Cylinder;
            }
            else if (
                type is eChartType.ConeBarClustered or eChartType.ConeBarStacked or eChartType.ConeBarStacked100 or eChartType.ConeCol or eChartType.ConeColClustered or eChartType.ConeColStacked or eChartType.ConeColStacked100)
            {
                Shape = eShape.Cone;
            }
            else if (
                type is eChartType.PyramidBarClustered or eChartType.PyramidBarStacked or eChartType.PyramidBarStacked100 or eChartType.PyramidCol or eChartType.PyramidColClustered or eChartType.PyramidColStacked or eChartType.PyramidColStacked100)
            {
                Shape = eShape.Pyramid;
            }
        }

        #endregion

        #region "Properties"

        readonly string _directionPath = "c:barDir/@val";

        /// <summary>
        /// Direction, Bar or columns
        /// </summary>
        public eDirection Direction
        {
            get => GetDirectionEnum(_chartXmlHelper.GetXmlNodeString(_directionPath));
            internal set => _chartXmlHelper.SetXmlNodeString(_directionPath, GetDirectionText(value));
        }

        readonly string _shapePath = "c:shape/@val";

        /// <summary>
        /// The shape of the bar/columns
        /// </summary>
        public eShape Shape
        {
            get => GetShapeEnum(_chartXmlHelper.GetXmlNodeString(_shapePath));
            internal set => _chartXmlHelper.SetXmlNodeString(_shapePath, GetShapeText(value));
        }

        ExcelChartDataLabel _DataLabel;

        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartDataLabel(NameSpaceManager, ChartNode);
                }

                return _DataLabel;
            }
        }

        readonly string _gapWidthPath = "c:gapWidth/@val";

        /// <summary>
        /// The size of the gap between two adjacent bars/columns
        /// </summary>
        public int GapWidth
        {
            get => _chartXmlHelper.GetXmlNodeInt(_gapWidthPath);
            set => _chartXmlHelper.SetXmlNodeString(_gapWidthPath, value.ToString(CultureInfo.InvariantCulture));
        }

        #endregion

        #region "Direction Enum Traslation"

        private string GetDirectionText(eDirection direction)
        {
            switch (direction)
            {
                case eDirection.Bar:
                    return "bar";
                default:
                    return "col";
            }
        }

        private eDirection GetDirectionEnum(string direction)
        {
            switch (direction)
            {
                case "bar":
                    return eDirection.Bar;
                default:
                    return eDirection.Column;
            }
        }

        #endregion

        #region "Shape Enum Translation"

        private string GetShapeText(eShape Shape)
        {
            switch (Shape)
            {
                case eShape.Box:
                    return "box";
                case eShape.Cone:
                    return "cone";
                case eShape.ConeToMax:
                    return "coneToMax";
                case eShape.Cylinder:
                    return "cylinder";
                case eShape.Pyramid:
                    return "pyramid";
                case eShape.PyramidToMax:
                    return "pyramidToMax";
                default:
                    return "box";
            }
        }

        private eShape GetShapeEnum(string text)
        {
            switch (text)
            {
                case "box":
                    return eShape.Box;
                case "cone":
                    return eShape.Cone;
                case "coneToMax":
                    return eShape.ConeToMax;
                case "cylinder":
                    return eShape.Cylinder;
                case "pyramid":
                    return eShape.Pyramid;
                case "pyramidToMax":
                    return eShape.PyramidToMax;
                default:
                    return eShape.Box;
            }
        }

        #endregion
    }
}