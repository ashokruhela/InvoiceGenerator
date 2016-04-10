using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace InvoiceGenerator
{
    internal class ExcelDataProvider
    {
        string filePath = string.Empty;
        public event EventHandler<string> UpdateProgress;
        Dictionary<int, string> columns = new Dictionary<int, string>();
        DataTable dtInvoice = new DataTable();
        public bool OpenExcelFile { get; set; }

        private void RaiseUpdateProgress(string currentValue)
        {
            if (UpdateProgress != null)
            {
                UpdateProgress(this, currentValue);
            }
        }

        public int GetMaxRows(string excelFilePath)
        {
            int rows = 0;
            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                Excel.Workbook currentWorkBook = excelApp.Workbooks.Open(excelFilePath);
                Excel.Worksheet currentSheet = currentWorkBook.Sheets[1];
                rows = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                Marshal.ReleaseComObject(currentWorkBook);
                Marshal.ReleaseComObject(currentSheet);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            return rows;
        }

        public Task<DataTable> GetExcelDataAsync(string excelFilePath)
        {
            return Task.Run<DataTable>(() =>
            {

                DataTable dtExcelData = new DataTable("Excel Data");
                columns.Clear();
                Excel.Application excelApp = null;
                try
                {
                    excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    Excel.Workbook currentWorkBook = excelApp.Workbooks.Open(excelFilePath);
                    Excel.Worksheet currentSheet = currentWorkBook.Sheets[1];
                    int lastRow = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    Excel.Range last = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    int lastColumn = currentSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    System.Array columnValues = (System.Array)currentSheet.get_Range("A1", last).Cells.Value;
                    for (int i = 0; i < lastColumn; i++)
                    {
                        DataColumn column = new DataColumn();
                        object val = columnValues.GetValue(1, i + 1);
                        column.Caption = val != null ? columnValues.GetValue(1, i + 1).ToString() : string.Empty;
                        dtExcelData.Columns.Add(column);
                        //Fill columns information
                        columns.Add(i, column.Caption);
                    }
                    //add extra columns to keep track which invoice to exclude
                    DataColumn skip = new DataColumn();
                    skip.Caption = Constants.Skip;
                    skip.DefaultValue = "NO";
                    columns.Add(columns.Count, skip.Caption);
                    dtExcelData.Columns.Add(skip);

                    if (!columns.ContainsValue(Constants.ExcluceInvoice))
                    {
                        DataColumn exclude = new DataColumn();
                        exclude.Caption = Constants.ExcluceInvoice;
                        exclude.DefaultValue = "NO";
                        columns.Add(columns.Count, exclude.Caption);
                        dtExcelData.Columns.Add(exclude);
                    }

                    for (int index = 2; index <= lastRow; index++)
                    {

                        DataRow newRow = dtExcelData.NewRow();
                        for (int i = 1; i <= lastColumn; i++)
                        {
                            object cellValue = columnValues.GetValue(index, i);
                            newRow[i - 1] = cellValue == null ? null : cellValue.ToString();
                        }
                        dtExcelData.Rows.Add(newRow);
                        string customerName = GetColumnValue(newRow, Constants.CustomerName);
                        if (customerName.Length > 0)
                            RaiseUpdateProgress(string.Format("Loading data for - {0}", customerName));
                    }
                    //delete rows to be excluded
                    List<DataRow> rowsToDelete = new List<DataRow>();
                    foreach (DataRow row in dtExcelData.Rows)
                    {
                        if (GetColumnValue(row, Constants.ExcluceInvoice).ToUpper() == "YES")
                            rowsToDelete.Add(row);
                    }
                    //delete excluded rows
                    foreach (DataRow row in rowsToDelete)
                    {
                        dtExcelData.Rows.Remove(row);
                    }
                }
                finally
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                    }
                }
                return dtExcelData;

            });


        }

        private string GetColumnValue(DataRow row, string columnCaption)
        {
            object cellValue = string.Empty;

            if (columns.ContainsValue(columnCaption))
            {
                var columnInfo = columns.FirstOrDefault(x => x.Value.ToUpper() == columnCaption.ToUpper());
                cellValue = row[columnInfo.Key];
            }

            return cellValue == null ? string.Empty : cellValue.ToString();
        }

        private Dictionary<string, string> GetColumnValue(DataRow row, List<string> columnCaptions)
        {
            Dictionary<string, string> columnValues = new Dictionary<string, string>();

            object cellValue = string.Empty;
            foreach (string columnCaption in columnCaptions)
            {
                if (columns.ContainsValue(columnCaption))
                {
                    var columnInfo = columns.FirstOrDefault(x => x.Value.ToUpper() == columnCaption.ToUpper());
                    cellValue = row[columnInfo.Key];
                    columnValues.Add(columnCaption, cellValue == null ? string.Empty : cellValue.ToString());
                }
            }

            return columnValues;
        }

        private void SetColumnValue(DataRow row, string columnCaption, string cellValue)
        {
            if (columns.ContainsValue(columnCaption))
            {
                var columnInfo = columns.FirstOrDefault(x => x.Value.ToUpper() == columnCaption.ToUpper());
                row[columnInfo.Key] = cellValue;
            }
        }

        public Task<bool> GenerateInvoice(DataTable dtTable)
        {
            return Task.Run(() =>
            {
                bool invoiceGenerated = true;
                Excel.Application excelApp = null;
                List<string> invoices = new List<string>();
                try
                {
                    ApplyFilterCondition(dtTable);
                    string tempValue = string.Empty;
                    foreach (DataRow row in dtTable.Rows)
                    {
                        bool isPvtLtd = true;
                        string custName = GetColumnValue(row, Constants.CustomerName);
                        if (GetColumnValue(row, Constants.Skip).ToUpper() == "YES" || custName.Length == 0)
                            continue;

                        string outputFilename = GetFileName(custName);
                        excelApp = new Excel.Application();
                        excelApp.Visible = false;
                        object misValue = System.Reflection.Missing.Value;
                        string supportEmail = @"info@shopon4u.com";
                        string companyTinNumber = string.Empty;
                        string website = "www.shopon4u.com";

                        #region Set Border
                        Excel.Workbook newWorkBook = excelApp.Workbooks.Add(misValue);
                        Excel.Worksheet newWorkSheet = (Excel.Worksheet)newWorkBook.Sheets.get_Item(1);
                        //fixed information
                        Excel.Range range = newWorkSheet.get_Range("A1", "K53");
                        range.Interior.Color = Excel.XlRgbColor.rgbWhite;
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("A1", "J28");
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("A9", "J9");
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("A10", "I16");
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("A19", "J20");
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("A17", "J18");
                        range.Merge();
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("A28", "J28");
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("G28", "J28");
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("A21", "B22");
                        range.Merge();
                        range.Value2 = Constants.ContestName;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("C21", "F22");
                        range.Merge();
                        range.Value2 = Constants.ContestName;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        Marshal.FinalReleaseComObject(range);

                        range = newWorkSheet.get_Range("G23", "J27");
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        Marshal.FinalReleaseComObject(range);



                        range = newWorkSheet.get_Range("G23", "I27");
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Set Width and height
                        newWorkSheet.Columns[1].ColumnWidth = 43.29;
                        newWorkSheet.Columns[2].ColumnWidth = .50;
                        newWorkSheet.Columns[3].ColumnWidth = 1.5;
                        newWorkSheet.Columns[4].ColumnWidth = 2;
                        newWorkSheet.Columns[5].ColumnWidth = 6;
                        newWorkSheet.Columns[6].ColumnWidth = 13;
                        newWorkSheet.Columns[7].ColumnWidth = 5;
                        newWorkSheet.Columns[8].ColumnWidth = 7;
                        newWorkSheet.Columns[9].ColumnWidth = 19.14;
                        newWorkSheet.Columns[10].ColumnWidth = 15;

                        newWorkSheet.Rows[1].RowHeight = 26;
                        newWorkSheet.Rows[10].RowHeight = 25;
                        newWorkSheet.Rows[11].RowHeight = 25;
                        newWorkSheet.Rows[12].RowHeight = 25;
                        newWorkSheet.Rows[13].RowHeight = 25;
                        newWorkSheet.Rows[14].RowHeight = 25;
                        newWorkSheet.Rows[15].RowHeight = 25;
                        newWorkSheet.Rows[16].RowHeight = 30;
                        #endregion

                        #region  Header
                        newWorkSheet.Cells[1, 1] = "Retail Invoice";
                        range = newWorkSheet.get_Range("A1", "J1");
                        range.Cells.Font.Size = 20;
                        range.Font.Bold = true;
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Merge();
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Company Logo
                        //company logo
                        string logoInfo = GetColumnValue(row, Constants.ModeOfPayment);
                        range = newWorkSheet.get_Range("A3");
                        try
                        {
                            if(Constants.IsStyleloby)
                            {
                                //styleloby
                                companyTinNumber = @"Company's TIN/VAT no :- 07606952779";
                                supportEmail = @"info@styleloby.com";
                                website = "www.styleloby.com";
                                isPvtLtd = true;
                                range.Value2 = "JMR Enterprises";
                                newWorkSheet.Shapes.AddPicture(Constants.StylelobyLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, 230, 40, 180, 50);
                            }
                            else if (logoInfo.Trim().ToUpper() == "NEFT (ICICI)")
                            {
                                //marketing
                                companyTinNumber = @"Company's TIN/VAT no :- 07706906681";
                                isPvtLtd = false;
                                range.Value2 = "Shopon Marketing";
                                newWorkSheet.Shapes.AddPicture(Constants.Shopon_M, MsoTriState.msoFalse, MsoTriState.msoCTrue, 210, 30, 220, 60);
                            }
                            else
                            {
                                //Private limited
                                companyTinNumber = @"Company's TIN/VAT no :- 07436920749";
                                isPvtLtd = true;
                                range.Value2 = "Shopon Marketing Pvt Ltd.";
                                newWorkSheet.Shapes.AddPicture(Constants.Shopon_P, MsoTriState.msoFalse, MsoTriState.msoCTrue, 230, 40, 180, 50);
                            }
                        }
                        catch (Exception)
                        {

                            if (System.Windows.Forms.MessageBox.Show("Failed to find the company logo images. Make sure files are present there.\n Would you like to generate without logo",
                                "Error occured", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Error) == System.Windows.Forms.DialogResult.No)
                            {
                                invoiceGenerated = false;
                                newWorkBook.Close(false, misValue, misValue);
                                Marshal.ReleaseComObject(newWorkSheet);
                                Marshal.ReleaseComObject(newWorkBook);
                                break;
                            }
                        }
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Company Name Information
                        //Telephone no
                        range = newWorkSheet.get_Range("A5");
                        range.Value2 = string.Format("Telephone: {0}",Constants.CustCareNumber);
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);

                        //tin number
                        range = newWorkSheet.get_Range("A6");
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        newWorkSheet.Hyperlinks.Add(range, @"mailto:" + supportEmail, Type.Missing, @"mailto:" + supportEmail, Type.Missing);
                        range.Value2 = companyTinNumber;
                        range.Cells.Font.Size = 11;
                        range.Font.Bold = true;
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        Marshal.FinalReleaseComObject(range);

                        //email id 
                        range = newWorkSheet.get_Range("A7");
                        newWorkSheet.Hyperlinks.Add(range, @"mailto:" + supportEmail, Type.Missing, @"mailto:" + supportEmail, Type.Missing);
                        range.Value2 = supportEmail;
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);

                        //Company name
                        range = newWorkSheet.get_Range("A4");
                        range.Value2 = "New Delhi, India";
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Date Added
                        //Order date
                        range = newWorkSheet.get_Range("I3");
                        range.Value2 = "Date Added:";
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        //Value
                        range = newWorkSheet.get_Range("J3");
                        tempValue = GetColumnValue(row, Constants.OrderDate);
                        DateTime orderDate = DateTime.Now;
                        DateTime.TryParse(tempValue, out orderDate);
                        range.Value2 = orderDate.ToString("dd/MM/yyyy");
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Reference No:
                        if (isPvtLtd)
                        {
                            range = newWorkSheet.get_Range("I4");
                            range.Value2 = "Reference No:";
                            range.Font.Bold = true;
                            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            range.Cells.Font.Size = 11;
                            Marshal.FinalReleaseComObject(range);
                            //value
                            range = newWorkSheet.get_Range("J4");
                            range.Value2 = GetColumnValue(row, Constants.RefNo);
                            range.Font.Bold = false;
                            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            range.Cells.Font.Size = 11;
                            Marshal.FinalReleaseComObject(range);
                        }
                        #endregion

                        #region Invoice number
                        //
                        range = newWorkSheet.get_Range(isPvtLtd ? "I5" : "I4");
                        range.Value2 = "Invoice Number:";
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range(isPvtLtd ? "J5" : "J4");
                        range.Value2 = GetInvoice(row);
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Order ID
                        //Order ID:
                        range = newWorkSheet.get_Range(isPvtLtd ? "I6" : "I5");
                        range.Value2 = "Order ID:";
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range(isPvtLtd ? "J6" : "J5");
                        range.NumberFormat = "@";
                        range.Font.Bold = false;
                        range.WrapText = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value = GetColumnValue(row, Constants.OrderID);
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Customer ID
                        //Customer ID
                        range = newWorkSheet.get_Range(isPvtLtd ? "I7" : "I6");
                        range.Value2 = "Customer ID:";
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range(isPvtLtd ? "J7" : "J6");
                        range.Value2 = GetColumnValue(row, Constants.CustomerID);
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Mode of Payment:
                        //Mode of Payment:
                        range = newWorkSheet.get_Range(isPvtLtd ? "I8" : "I7");
                        range.Value2 = "Mode of Payment:";
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range(isPvtLtd ? "J8" : "J7");
                        range.Value2 = "Prepaid";
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Billing Address
                        // Billing Address
                        range = newWorkSheet.get_Range("A9", "D9");
                        range.Merge();
                        range.Font.Bold = true;
                        range.Value2 = "Billing Address";
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("A10", "D16");
                        range.Merge();
                        range.Font.Bold = false;
                        range.Value2 = GetBillingAddress(row);
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        range.WrapText = true;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Shipping address
                        range = newWorkSheet.get_Range("E9", "J9");
                        range.Merge();
                        range.Font.Bold = true;
                        range.Value2 = "Shipping Address";
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("E10", "J16");
                        range.Merge();
                        range.Font.Bold = false;
                        range.WrapText = true;
                        range.Value2 = GetShippingAddress(row); ;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Product & Model
                        // Product
                        range = newWorkSheet.get_Range("A19", "B20");
                        range.Merge();
                        range.Font.Bold = true;
                        range.Value2 = "Product";
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);
                        // value
                        range = newWorkSheet.get_Range("A23", "F27");
                        range.Merge();
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        range.Font.Bold = true;
                        range.WrapText = true;
                        range.Value2 = GetProducts(row);
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Cells.Font.Size = 11;
                        Marshal.FinalReleaseComObject(range);

                        //Model
                        range = newWorkSheet.get_Range("C19", "F20");
                        range.Merge();
                        range.Font.Bold = true;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Value2 = "Model";
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region QTY
                        //QTY
                        range = newWorkSheet.get_Range("G19", "G20");
                        range.Merge();
                        range.Font.Bold = true;
                        range.Value2 = "QTY";
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("G21", "G22");
                        range.Merge();
                        range.Font.Bold = false;
                        range.Value2 = GetColumnValue(row, Constants.QTY);
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region MRP
                        //MRP
                        range = newWorkSheet.get_Range("H19", "H20");
                        range.Merge();
                        range.Font.Bold = true;
                        range.Value2 = "MRP";
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("H21", "H22");
                        range.Merge();
                        range.Font.Bold = false;
                        range.Value2 = GetColumnValue(row, Constants.MRP);
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Gross Value
                        //Gross Value
                        range = newWorkSheet.get_Range("I19", "I20");
                        range.Merge();
                        range.Font.Bold = true;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Value2 = "Gross Value";
                        Marshal.FinalReleaseComObject(range);
                        //value
                        string tempprice = GetColumnValue(row, Constants.OrderValue);
                        bool hasPaisa = !tempprice.Contains(".");
                        string price = !hasPaisa ? tempprice : string.Format("{0}.00", tempprice);
                        range = newWorkSheet.get_Range("I21", "I22");
                        range.Merge();
                        range.NumberFormat = "0.00";
                        range.Font.Bold = false;
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Value2 = price;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Net Amount
                        //Net Amount
                        range = newWorkSheet.get_Range("J19", "J20");
                        range.Merge();
                        range.Font.Bold = true;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Value2 = "Net Amount";
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("J21", "J22");
                        range.Merge();
                        range.NumberFormat = "0.00";
                        range.Font.Bold = false;
                        range.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        range.Value2 = price;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Sub Total
                        //Sub Total:
                        range = newWorkSheet.get_Range("G23", "I23");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Sub Total:";
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("J23");
                        range.NumberFormat = "0.00";
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        range.Value2 = price;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Flat Shipping rate
                        //Flat Shipping rate:
                        range = newWorkSheet.get_Range("G24", "I24");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Flat Shipping rate:";
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("J24");
                        range.NumberFormat = "0.00";
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        double shippingCost = 0.0;
                        double.TryParse(GetColumnValue(row, Constants.Shipping), out shippingCost);
                        range.Value2 = shippingCost;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Service Tax
                        //Service Tax
                        range = newWorkSheet.get_Range("G25", "I25");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Service Tax:";
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("J25");
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        range.NumberFormat = "0.00";
                        range.Value2 = "0.00";
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Invoice Amount
                        //Invoice Amount- Incl. of all Taxes
                        range = newWorkSheet.get_Range("G26", "I26");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Invoice Amount- Incl. of all Taxes";
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("J26");
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        range.NumberFormat = "0.00";
                        range.Value2 = price;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Handling Charges
                        //Handling Charges:
                        range = newWorkSheet.get_Range("G27", "I27");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Handling Charges:";
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("J27");
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        range.NumberFormat = "0.00";
                        range.Value2 = "0.00";
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Amount Payable
                        //Amount Payable:
                        range = newWorkSheet.get_Range("G28", "I28");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Amount Payable:";
                        Marshal.FinalReleaseComObject(range);
                        //value
                        range = newWorkSheet.get_Range("J28");
                        range.NumberFormat = "0.00";
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        double finalPrice = Convert.ToDouble(tempprice) + shippingCost;
                        range.Value2 = finalPrice;
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        #region Bottom Fixed information
                        //Amount to be Collected :- 
                        range = newWorkSheet.get_Range("A32");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Amount to be Collected :- ";
                        Marshal.FinalReleaseComObject(range);

                        //Amount in words
                        range = newWorkSheet.get_Range("B32", "J32");
                        range.Merge();
                        range.Font.Bold = true;
                        //string fprice = !hasPaisa ? finalPrice : string.Format("{0}.00", finalPrice);
                        //string words = hasPaisa ? HelpUtil.ToWords(Convert.ToDecimal(price)) : HelpUtil.ToWords(Convert.ToDecimal(tempprice));
                        string words = HelpUtil.ToWords(Convert.ToDecimal(finalPrice));
                        range.Value2 = string.Format("{0} ONLY", words.ToUpper());
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A33");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Additional Comments: ";
                        Marshal.FinalReleaseComObject(range);

                        //Term and Conditions :
                        range = newWorkSheet.get_Range("A37");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "Term and Conditions : ";
                        Marshal.FinalReleaseComObject(range);

                        //(1) Refunds will be made as per our refund policy.
                        range = newWorkSheet.get_Range("A38");
                        range.Merge();
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "(1) Refunds will be made as per our refund policy.";
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A39");
                        range.Merge();
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "(2) VAT/CST is applicable on above amount is: - Rs 000.00/- ";
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A40");
                        range.Merge();
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = string.Format("(3) In Case of any queries, please call our customer care on: {0} or email: {1}", Constants.CustCareNumber, supportEmail);
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A41");
                        range.Merge();
                        range.Font.Bold = false;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "(4) All disputes are subject to the exclusive jurisdiction of competent courts and forums in Delhi/New Delhi only.";
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A44", "J44");
                        range.Merge();
                        range.Font.Bold = true;
                        range.Value2 = "Visit us At : " + website;

                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A45", "J45");
                        range.Merge();
                        range.Font.Bold = true;
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Value2 = "This is a computer generated invoice. No signature required.";
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A50");
                        range.Merge();
                        range.Font.Bold = true;
                        range.Value2 = "Thank you";
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A51");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = "for shopping with us.";
                        Marshal.FinalReleaseComObject(range);

                        //Additional Comments:
                        range = newWorkSheet.get_Range("A52");
                        range.Merge();
                        range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        range.Value2 = website;
                        newWorkSheet.Hyperlinks.Add(range, @"http://" + website, Type.Missing, @"http://" + website, Type.Missing);
                        Marshal.FinalReleaseComObject(range);
                        #endregion

                        newWorkSheet.PageSetup.Zoom = false;
                        newWorkSheet.PageSetup.FitToPagesWide = 1;
                        newWorkSheet.PageSetup.FitToPagesTall = 1;
                        newWorkBook.SaveAs(outputFilename, Excel.XlFileFormat.xlOpenXMLWorkbook,
                            System.Reflection.Missing.Value, Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                            Excel.XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value,
                            Missing.Value);
                        newWorkBook.Close(true, misValue, misValue);
                        Marshal.ReleaseComObject(newWorkSheet);
                        Marshal.ReleaseComObject(newWorkBook);
                        if (OpenExcelFile)
                        {
                            string[] temp = outputFilename.Split(new char[] { '\\' });
                            string file = temp[temp.Count() - 1];
                            string currentDate = DateTime.Today.ToString(Constants.FolderNameFormat);
                            var files = Directory.GetFiles(Path.Combine(Constants.OutputFilePath, currentDate), string.Format("{0}.*", file));
                            if (files.Count() > 0)
                            {
                                invoices.Add(files[0]);
                            }
                        }
                        RaiseUpdateProgress(string.Format("Generated invoice for  {0}", custName));
                    }
                }
                finally
                {
                    if(excelApp == null)
                    {
                        System.Windows.Forms.MessageBox.Show("No data to process. Either customer name could be empty, column name is invalid or skip is YES in sheet. Make sure that valid data is present.", 
                            "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                    else
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                        try
                        {
                            foreach (string invoice in invoices)
                            {
                                System.Diagnostics.Process.Start(invoice);
                            }

                        }
                        catch { }
                    }
                    
                }
                return invoiceGenerated;
            });
        }

        public void ApplyFilterCondition(DataTable dtRawData)
        {
            //condition 1 - if cust id, order number, invoice no & ref no are same then 
            //there would be only one invoice and amount will be added
            FilterCondition(dtRawData, 1);

            //condition 2 - if Only customer id & invoice no are same then the amount will be added but both invoice no has to be displayed
            FilterCondition(dtRawData, 2);
        }

        private void FilterCondition(DataTable dtRawData, int condition)
        {

            List<string> captionNames = new List<string>(new string[] { Constants.CustomerID, Constants.OrderID, Constants.InvoiceNo, Constants.RefNo, Constants.Products });
            foreach (DataRow row in dtRawData.Rows)
            {
                Dictionary<string, string> mainRowValues = GetColumnValue(row, captionNames);
                string custID = mainRowValues[Constants.CustomerID];
                if (GetColumnValue(row, Constants.Skip).ToUpper() == "YES" || custID.Length == 0)
                    continue;

                string orderID = mainRowValues[Constants.OrderID];
                string invoiceNo = mainRowValues[Constants.InvoiceNo];
                string refNo = mainRowValues[Constants.RefNo];
                string products = string.Empty;
                string newProducts = products = mainRowValues[Constants.Products];
                double totalPrice = Convert.ToDouble(GetColumnValue(row, Constants.OrderValue));
                string newOrderID = orderID;
                foreach (DataRow otherRow in dtRawData.Rows)
                {
                    if (GetColumnValue(otherRow, Constants.Skip).ToUpper() == "YES")
                        continue;
                    if (!row.Equals(otherRow))
                    {
                        Dictionary<string, string> otherRowValues = GetColumnValue(otherRow, captionNames);
                        string othercustID = otherRowValues[Constants.CustomerID];
                        string otherorderID = otherRowValues[Constants.OrderID];
                        string otherinvoiceNo = otherRowValues[Constants.InvoiceNo];
                        string otherrefNo = otherRowValues[Constants.RefNo];
                        string otherProducts = otherRowValues[Constants.Products];

                        if (condition == 1 && custID == othercustID && orderID == otherorderID && invoiceNo == otherinvoiceNo && refNo == otherrefNo)
                        {
                            newProducts = string.Join("+", newProducts, otherProducts);
                            totalPrice += Convert.ToDouble(GetColumnValue(otherRow, Constants.OrderValue));
                            SetColumnValue(row, Constants.Products, newProducts);
                            SetColumnValue(otherRow, Constants.Skip, "Yes");
                            SetColumnValue(row, Constants.OrderValue, totalPrice.ToString());
                        }
                        else if (condition == 2 && custID == othercustID && invoiceNo == otherinvoiceNo)
                        {
                            newProducts = string.Join("+", newProducts, otherProducts);
                            totalPrice += Convert.ToDouble(GetColumnValue(otherRow, Constants.OrderValue));
                            string lastDigits = otherorderID.Substring(otherorderID.Length - 5, 5);
                            newOrderID = string.Join("/", newOrderID, lastDigits);
                            SetColumnValue(otherRow, Constants.Skip, "Yes");
                            SetColumnValue(row, Constants.Products, newProducts);
                            SetColumnValue(row, Constants.OrderValue, totalPrice.ToString());
                            SetColumnValue(row, Constants.OrderID, newOrderID);
                        }
                    }
                }
            }
        }

        private string GetShippingAddress(DataRow row)
        {
            string gender = GetColumnValue(row, Constants.Gender);
            if (gender.ToUpper() == "F")
                gender = "Mrs.";
            else if (gender.ToUpper() == "U")
                gender = "Miss";
            else
                gender = "Mr.";
            string customerName = GetColumnValue(row, Constants.CustomerName);
            string address = GetColumnValue(row, Constants.Address);
            string city = GetColumnValue(row, Constants.City);
            string state = GetColumnValue(row, Constants.State);
            string pincode = GetColumnValue(row, Constants.Pincode);
            string emailID = GetColumnValue(row, Constants.EmailID);
            string phone = GetColumnValue(row, Constants.Phone);
            string alternateNo = GetColumnValue(row, Constants.AlternameNumber);
            if (alternateNo.Length > 0 && alternateNo.Trim() != "-")
            {
                phone = string.Join("/", phone, alternateNo);
            }
            string billingAddress = string.Format("{0} {1}\n\n{2}\n\n{3}\n{4} Pin {5}\nIndia\n{6}\n+91{7}", gender, customerName, address, city, state, pincode, emailID, phone);
            return billingAddress;
        }

        private string GetBillingAddress(DataRow row)
        {
            return GetShippingAddress(row);
        }

        private string GetInvoice(DataRow row)
        {
            string invoiceNo = string.Empty;
            invoiceNo = GetColumnValue(row, Constants.InvoiceNo);
            if (invoiceNo.Length == 0)
                invoiceNo = GetColumnValue(row, Constants.RefNo);
            return invoiceNo;
        }

        private string GetFileName(string customerName)
        {
            customerName = customerName.Replace(".", string.Empty);
            string currentDate = DateTime.Today.ToString(Constants.FolderNameFormat);
            string fileName = Path.Combine(Constants.OutputFilePath, currentDate, customerName);
            if (!Directory.Exists(Path.Combine(Constants.OutputFilePath, currentDate)))
            {
                Directory.CreateDirectory(Path.Combine(Constants.OutputFilePath, currentDate));
            }
            var files = Directory.GetFiles(Path.Combine(Constants.OutputFilePath, currentDate), string.Format("{0}*", customerName));
            fileName = files.Count() == 0 ? fileName : string.Format("{0}_{1}", fileName, (files.Count() + 1).ToString());

            if (File.Exists(fileName))
            {
                Random randomNumber = new Random();
                fileName = string.Format("{0}_{1}", fileName, randomNumber.Next(50, 100));
            }

            return fileName;
        }

        private string GetProducts(DataRow row)
        {
            var temp = GetColumnValue(row, Constants.Products).Split(new char[] { '+' });
            StringBuilder products = new StringBuilder();
            int productCount = temp.Count();

            foreach (string product in temp)
            {
                products.AppendLine(product.Trim());
            }
            if (productCount > 4)
            {
                products = products.Replace("\n", ", ");
                //Remove last comma
                products.Remove(products.Length - 3, 2);
            }
            return products.ToString();
        }
    }
}
