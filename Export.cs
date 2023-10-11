using CrowdSafe.Models.DataModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;

namespace CrowdSafe.Helpers
{
    public static class Export
    {
        public static string ExcelContentType
        {
            get
            { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; }
        }
        public static DataTable ListToDataTable<T>(List<T> data, string[] headers)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable dataTable = new DataTable();

            for (int i = 0; i < headers.Count(); i++)
            {
                PropertyDescriptor property = properties[i];
                dataTable.Columns.Add(headers[i]);
            }

            object[] values = new object[properties.Count];
            if (data.Count() > 0)
            {
                foreach (T item in data)
                {
                    for (int i = 0; i < values.Length; i++)
                    {
                        values[i] = properties[i].GetValue(item);
                    }

                    dataTable.Rows.Add(values);
                }
            }
            else
            {
                for (int i = 0; i < headers.Count(); i++)
                {
                    values[i] = " ";
                }

                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
        public static byte[] ToExcel(DataTable dataTable, string heading = "", bool showSrNo = false, string headerLine = "", bool TotalColumn = false, params string[] columnsToTake)
        {

            byte[] result = null;
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(String.Format("{0} Data", heading));
                int startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 2;
                int rowCount = (dataTable.Rows.Count + 2);
                int columnCount = (showSrNo) ? (dataTable.Columns.Count + 1) : dataTable.Columns.Count;
                if (showSrNo)
                {
                    DataColumn dataColumn = dataTable.Columns.Add("#", typeof(int));
                    dataColumn.SetOrdinal(0);
                    int index = 1;
                    foreach (DataRow item in dataTable.Rows)
                    {
                        if (TotalColumn)
                        {
                            if (!(index == dataTable.Rows.Count))
                            {
                                item[0] = index;
                                index++;
                            }
                        }
                        else
                        {
                            item[0] = index;
                            index++;
                        }
                    }
                }


                // add the content into the Excel file  
                workSheet.Cells["A" + startRowFrom].LoadFromDataTable(dataTable, true);
                int columnIndex = 1;
                foreach (DataColumn column in dataTable.Columns)
                {
                    workSheet.Column(columnIndex);
                    columnIndex++;
                }

                // format header - bold, yellow on black  
                using (ExcelRange r = workSheet.Cells[startRowFrom, 1, startRowFrom, dataTable.Columns.Count])
                {
                    r.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    r.Style.Font.Bold = true;
                    r.Style.Font.Size = 13;
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#398E92"));
                }

                // format cells - add borders  
                using (ExcelRange r = workSheet.Cells[startRowFrom + 1, 1, startRowFrom + dataTable.Rows.Count, dataTable.Columns.Count])
                {
                    r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                }
                //// removed ignored columns  
                //for (int i = dataTable.Columns.Count - 1; i >= 0; i--)
                //{

                //    if (i == 0 && showSrNo)
                //    {
                //        continue;
                //    }
                //    if (!columnsToTake.Contains(dataTable.Columns[i].ColumnName))
                //    {
                //        workSheet.DeleteColumn(i + 1);
                //    }
                //}
                if (!String.IsNullOrEmpty(heading))
                {
                    workSheet.Cells["A1"].Value = heading;
                    workSheet.Cells["A1"].Style.Font.Size = 15;
                    workSheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    workSheet.Cells["A1"].Style.Font.Bold = true;
                    workSheet.Cells["A1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#398E92"));
                    workSheet.Cells["A1"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells["A1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells["A1"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells["A1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells["A1"].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    workSheet.Cells["A1"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    workSheet.Cells["A1"].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    workSheet.Cells["A1"].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    if (headerLine == null || headerLine.ToString().Trim() == "" && columnsToTake.Length <= 26)
                    {
                        workSheet.Cells["A1:" + Convert.ToChar(65 + columnsToTake.Length - ((showSrNo) ? 0 : 1)) + "1"].Merge = true;
                    }
                    else if (columnsToTake.Length > 26)
                    {
                        var collengthA1 = columnsToTake.Length - 26;
                        workSheet.Cells["A1:A" + Convert.ToChar(65 + collengthA1 - ((showSrNo) ? 0 : 1)) + "1"].Merge = true;
                    }
                    else
                    {
                        workSheet.Cells["B1"].Value = headerLine;
                        workSheet.Cells["B1"].Style.Font.Size = 15;
                        workSheet.Cells["B1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["B1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        workSheet.Cells["B1"].Style.Font.Bold = true;
                        workSheet.Cells["B1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells["B1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#398E92"));
                        workSheet.Cells["B1"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["B1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["B1"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["B1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["B1"].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["B1"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["B1"].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["B1"].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["B1:" + Convert.ToChar(65 + columnsToTake.Length - ((showSrNo) ? 0 : 1)) + "1"].Merge = true;
                        if (columnsToTake.Length <= 26)
                        {
                            workSheet.Cells["B1:" + Convert.ToChar(65 + columnsToTake.Length - ((showSrNo) ? 0 : 1)) + "1"].Merge = true;
                        }
                        else
                        {
                            var collength = columnsToTake.Length - 26;
                            workSheet.Cells["B1:A" + Convert.ToChar(65 + collength - ((showSrNo) ? 0 : 1)) + "1"].Merge = true;
                        }
                    }
                    if (TotalColumn)
                    {
                        for (int columnNumber = 0; columnNumber < columnCount; columnNumber++)
                        {
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Font.Size = 13;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Font.Bold = true;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#ffee00"));
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            workSheet.Cells[Convert.ToChar(65 + columnNumber - ((showSrNo) ? 0 : 1)) + rowCount.ToString()].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        }

                    }
                    workSheet.InsertColumn(1, 1);
                    workSheet.InsertRow(1, 1);
                    workSheet.Column(1).Width = 5;
                }
                result = package.GetAsByteArray();
            }

            return result;
        }

        public static byte[] ToExcel<T>(List<T> data, string[] ColumnsHeaders, string Heading = "", bool showSlno = false)
        {
            return ToExcel(ListToDataTable<T>(data, ColumnsHeaders), Heading, showSlno, "", false, ColumnsHeaders);
        }

        public static byte[] StaffingSheet(List<StaffingSheet> StaffingSheetList, string[] ColumnsHeaders, params string[] columnsToTake)
        {
            byte[] result = null;
            var GroupStaffingSheet = StaffingSheetList.OrderBy(x => x.Position).GroupBy(x => x.Position, (Position, StewardsList) => new { StewardPosition = Position, Stewards = StewardsList.ToList() }).ToList();
            using (ExcelPackage package = new ExcelPackage())
            {
                int SheetNo = 0;
                foreach (var StewardsList in GroupStaffingSheet)
                {
                    string heading = "Staffing Sheet";// StewardsList.StewardPosition.Replace(" ", "_");
                    DataTable dataTable = ListToDataTable<StaffingSheet>(StewardsList.Stewards, ColumnsHeaders);
                    ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(String.Format("{0}", "Sheet " + (++SheetNo)));
                    int startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 2;
                    int rowCount = (dataTable.Rows.Count + 2);

                    DataColumn dataColumn = dataTable.Columns.Add("#", typeof(int));
                    dataColumn.SetOrdinal(0);
                    int index = 1;
                    foreach (DataRow item in dataTable.Rows)
                    {
                        item[0] = index;
                        index++;
                    }



                    // add the content into the Excel file  
                    workSheet.Cells["A" + startRowFrom].LoadFromDataTable(dataTable, true);
                    int columnIndex = 1;
                   // workSheet.Column(columnIndex).AutoFit();
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        workSheet.Column(columnIndex).Width=20;//.AutoFit();
                        columnIndex++;
                    }

                    // format header - bold, yellow on black  
                    using (ExcelRange r = workSheet.Cells[startRowFrom, 1, startRowFrom, dataTable.Columns.Count])
                    {
                        r.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        r.Style.Font.Bold = true;
                        r.Style.Font.Size = 13;
                        r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#398E92"));
                    }

                    // format cells - add borders  
                    using (ExcelRange r = workSheet.Cells[startRowFrom, 1, startRowFrom + dataTable.Rows.Count, dataTable.Columns.Count])
                    {
                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    }

                    if (!String.IsNullOrEmpty(heading))
                    {
                        workSheet.Cells["A1"].Value = heading;
                        workSheet.Cells["A1"].Style.Font.Size = 15;
                        workSheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        workSheet.Cells["A1"].Style.Font.Bold = true;
                        workSheet.Cells["A1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#398E92"));
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Merge = true;
                        workSheet.InsertColumn(1, 1);
                        workSheet.InsertRow(1, 1);
                        workSheet.Column(1).Width = 5;
                    }
                }
                result = package.GetAsByteArray();
            }

            return result;
        }

        public static byte[] PreOderStaffingSheet(List<PreOrderStaffingSheet> StaffingSheetList, string[] ColumnsHeaders, params string[] columnsToTake)
        {
            byte[] result = null;
            var GroupStaffingSheet = StaffingSheetList.OrderBy(x => x.Position).GroupBy(x => x.Position, (Position, StewardsList) => new { StewardPosition = Position, Stewards = StewardsList.ToList() }).ToList();
            using (ExcelPackage package = new ExcelPackage())
            {
                int SheetNo = 0;
                foreach (var StewardsList in GroupStaffingSheet)
                {

                    string heading = "Staffing Sheet";// StewardsList.StewardPosition.Replace(" ", "_"); ;
                    DataTable dataTable = ListToDataTable<PreOrderStaffingSheet>(StewardsList.Stewards, ColumnsHeaders);
                    ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(String.Format("{0}", "Sheet " + (++SheetNo)));
                    int startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 2;
                    int rowCount = (dataTable.Rows.Count + 2);

                    DataColumn dataColumn = dataTable.Columns.Add("#", typeof(int));
                    dataColumn.SetOrdinal(0);
                    int index = 1;
                    foreach (DataRow item in dataTable.Rows)
                    {
                        item[0] = index;
                        index++;
                    }



                    // add the content into the Excel file  
                    workSheet.Cells["A" + startRowFrom].LoadFromDataTable(dataTable, true);
                    int columnIndex = 1;
                   
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        workSheet.Column(columnIndex).Width=20;///AutoFit();
                        columnIndex++;
                    }

                    // format header - bold, yellow on black  
                    using (ExcelRange r = workSheet.Cells[startRowFrom, 1, startRowFrom, dataTable.Columns.Count])
                    {
                        r.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        r.Style.Font.Bold = true;
                        r.Style.Font.Size = 13;
                        r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#398E92"));
                    }

                    // format cells - add borders  
                    using (ExcelRange r = workSheet.Cells[startRowFrom, 1, startRowFrom + dataTable.Rows.Count, dataTable.Columns.Count])
                    {
                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    }

                    if (!String.IsNullOrEmpty(heading))
                    {
                        workSheet.Cells["A1"].Value = heading;
                        workSheet.Cells["A1"].Style.Font.Size = 15;
                        workSheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        workSheet.Cells["A1"].Style.Font.Bold = true;
                        workSheet.Cells["A1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#398E92"));
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        workSheet.Cells["A1:" + Convert.ToChar(65 + ColumnsHeaders.Length) + "1"].Merge = true;
                        workSheet.InsertColumn(1, 1);
                        workSheet.InsertRow(1, 1);
                        workSheet.Column(1).Width = 5;
                    }
                }
                result = package.GetAsByteArray();
            }

            return result;
        }

    }
}


