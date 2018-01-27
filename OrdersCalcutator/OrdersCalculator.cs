using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;

namespace OrdersCalcutator
{
    public static class OrdersCalculator
    {
        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            var stringTablePart = document.WorkbookPart.SharedStringTablePart;
            var value = cell.CellValue?.InnerXml;
            if (value == null)
                return null;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
            }

            return value;
        }

        public static void CalculateOrders(string[] files, DateTime startDate, DateTime finishDate)
        {
            var finishOrders = new List<Order>();

            foreach (var file in files)
            {
                using (var doc = SpreadsheetDocument.Open(file, false))
                {
                    var sheets = doc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                    var relationshipId = sheets.First().Id.Value;
                    var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(relationshipId);
                    var workSheet = worksheetPart.Worksheet;
                    var sheetData = workSheet.GetFirstChild<SheetData>();
                    var rows = sheetData.Descendants<Row>()
                                .Skip(1)
                                .ToArray();
                    var orders = ParseData(rows, doc);
                    finishOrders.AddRange(orders);
                }
            }

            var filterOrders = FilterOrders(finishOrders, startDate, finishDate);
            CreateExcelFile(filterOrders);
        }

        private static List<Order> FilterOrders(List<Order> orders, DateTime startDate, DateTime finishDate)
        {
            var zeroOrders = GetZeroOrders(orders); 
            return orders
                        .Where(order => zeroOrders.Contains(order.CompanyName + order.MainContact)
                                        && !string.IsNullOrEmpty(order.CompanyName) && !string.IsNullOrEmpty(order.MainContact)
                                        && order.CreateDate >= startDate && order.CreateDate <= finishDate
                                        && Math.Abs(order.Budget) > 0.0001)
                        .Distinct()
                        .ToList();
        }

        private static HashSet<string> GetZeroOrders(List<Order> orders)
        {
            var zeroOrdersDict = new Dictionary<string, DateTime>();
            foreach (var order in orders.Where(order => Math.Abs(order.Budget) < 0.0001))
            {
                if (!zeroOrdersDict.ContainsKey(order.CompanyName + order.MainContact))
                    zeroOrdersDict.Add(order.CompanyName + order.MainContact, order.CreateDate);
                else if (zeroOrdersDict[order.CompanyName + order.MainContact] > order.CreateDate)
                    zeroOrdersDict[order.CompanyName + order.MainContact] = order.CreateDate;
            }

            foreach (var order in orders)
            {
                if (zeroOrdersDict.ContainsKey(order.CompanyName + order.MainContact)
                    && zeroOrdersDict[order.CompanyName + order.MainContact] > order.CreateDate)
                {
                    zeroOrdersDict.Remove(order.CompanyName + order.MainContact);
                }
            }

            return new HashSet<string>(zeroOrdersDict.Keys);
        }

        private static void CreateExcelFile(List<Order> orders)
        {
            var excelApp = new Excel.Application
            {
                Visible = false,
                ScreenUpdating = false
            };
            var workBook = excelApp.Workbooks.Add();
            var workSheet = (Excel.Worksheet)workBook.Worksheets.Item[1];
            var counter = 2;
            NamedColumns(workSheet);

            foreach (var order in orders)
            {
                workSheet.Cells[counter, 1] = order.TradeName;
                workSheet.Cells[counter, 2] = order.CompanyName;
                workSheet.Cells[counter, 3] = order.MainContact;
                workSheet.Cells[counter, 4] = order.CompanyContact;
                workSheet.Cells[counter, 5] = order.Responsible;
                workSheet.Cells[counter, 6] = order.TransactionStage;
                workSheet.Cells[counter, 7] = order.Budget.ToString(CultureInfo.CurrentCulture);
                workSheet.Cells[counter, 8] = order.CreateDate.ToShortDateString();
                workSheet.Cells[counter, 9] = order.Creator;
                workSheet.Cells[counter, 10] = order.ChangeDate.ToShortDateString();

                workSheet.Cells[counter, 11] = order.DeliveryAddress;
                workSheet.Cells[counter, 12] = order.OrderDate == null ? "" : order.OrderDate.Value.ToShortDateString();

                workSheet.Cells[counter, 13] = order.WorkingEmail;
                workSheet.Cells[counter, 14] = order.PrivateEmail;

                workSheet.Cells[counter, 15] = order.WorkingPhone;
                workSheet.Cells[counter, 16] = order.MobilePhone;
                counter++;
            }
            var sum = orders.Where(order => order.TransactionStage == "Выполнен" 
                                            || order.TransactionStage == "Оплачен" 
                                            || order.TransactionStage == "Успешно реализовано")
                            .Sum(order => order.Budget);
            workSheet.Cells[counter, 1] = "Итого:";
            workSheet.Cells[counter, 2] = $"{sum} руб";

            var filename = "Result_Calc_orders.xls";
            var misValue = System.Reflection.Missing.Value;
            var path = Path.Combine(Directory.GetCurrentDirectory(), filename);
            workBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            workBook.Close(true, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            excelApp.Quit();
        }

        private static void NamedColumns(Excel._Worksheet worksheet)
        {
            worksheet.Columns[1].ColumnWidth = 15;
            worksheet.Columns[2].ColumnWidth = 20;
            worksheet.Columns[3].ColumnWidth = 20;
            worksheet.Columns[4].ColumnWidth = 20;

            worksheet.Cells[1, 1] = "Название сделки";
            worksheet.Cells[1, 2] = "Компания";
            worksheet.Cells[1, 3] = "Основной контакт";
            worksheet.Cells[1, 4] = "Компания контакта";
            worksheet.Cells[1, 5] = "Отвественный";
            worksheet.Cells[1, 6] = "Этап сделки";
            worksheet.Cells[1, 7] = "Бюджет";
            worksheet.Cells[1, 8] = "Дата создания";
            worksheet.Cells[1, 9] = "Кем создана";
            worksheet.Cells[1, 10] = "Дата изменения";
            worksheet.Cells[1, 11] = "Адрес доставки";
            worksheet.Cells[1, 12] = "Дата заказа";
            worksheet.Cells[1, 13] = "Рабочий Email (контакт)";
            worksheet.Cells[1, 14] = "Личный Email (контакт)";
            worksheet.Cells[1, 15] = "Рабочий телефон (контакт)";
            worksheet.Cells[1, 16] = "Мобильный телефон (контакт)";
        }


        private static List<Order> ParseData(Row[] rows, SpreadsheetDocument doc)
        {
            var orders = new List<Order>();
            foreach (var row in rows)
            {
                var order = new Order
                {
                    TradeName = GetCellValue(doc, row.Elements<Cell>().ElementAt(0)),
                    CompanyName = GetCellValue(doc, row.Elements<Cell>().ElementAt(1)),
                    MainContact = GetCellValue(doc, row.Elements<Cell>().ElementAt(2)),
                    CompanyContact = GetCellValue(doc, row.Elements<Cell>().ElementAt(3)),
                    Responsible = GetCellValue(doc, row.Elements<Cell>().ElementAt(4)),
                    TransactionStage = GetCellValue(doc, row.Elements<Cell>().ElementAt(5)),
                    Budget = string.IsNullOrEmpty(GetCellValue(doc, row.Elements<Cell>().ElementAt(6))) ? 0 : double.Parse(GetCellValue(doc, row.Elements<Cell>().ElementAt(6))),
                    CreateDate = DateTime.Parse(GetCellValue(doc, row.Elements<Cell>().ElementAt(7))).Date,
                    Creator = GetCellValue(doc, row.Elements<Cell>().ElementAt(8)),
                    ChangeDate = DateTime.Parse(GetCellValue(doc, row.Elements<Cell>().ElementAt(9))).Date,

                    DeliveryAddress = GetCellValue(doc, row.Elements<Cell>().ElementAt(15)),
                    OrderDate = string.IsNullOrEmpty(GetCellValue(doc, row.Elements<Cell>().ElementAt(16))) ? (DateTime?) null
                                       : DateTime.Parse(GetCellValue(doc, row.Elements<Cell>().ElementAt(16))).Date,

                    WorkingEmail = GetCellValue(doc, row.Elements<Cell>().ElementAt(28)),
                    PrivateEmail = GetCellValue(doc, row.Elements<Cell>().ElementAt(29)),

                    WorkingPhone = GetCellValue(doc, row.Elements<Cell>().ElementAt(31)),
                    MobilePhone = GetCellValue(doc, row.Elements<Cell>().ElementAt(33))
                };

                orders.Add(order);

            //}
            //for (var i = 2; i <= cells.Rows.Count; i++)
            //{
                //var cells = workSheet.UsedRange;
                //var orders = new List<Order>();
                //for (var i = 2; i <= cells.Rows.Count; i++)
                //{
                //    var order = new Order
                //    {
                //        TradeName = cells.Cells[i, 1].Value2?.ToString(),
                //        CompanyName = cells.Cells[i, 2].Value2?.ToString(),
                //        MainContact = cells.Cells[i, 3].Value2?.ToString(),
                //        CompanyContact = cells.Cells[i, 4].Value2?.ToString(),
                //        Responsible = cells.Cells[i, 5].Value2?.ToString(),
                //        TransactionStage = cells.Cells[i, 6].Value2?.ToString(),
                //        Budget = double.Parse(cells.Cells[i, 7].Value2?.ToString() ?? 0),
                //        CreateDate = DateTime.Parse(cells.Cells[i, 8].Value2).Date,
                //        Creator = cells.Cells[i, 9].Value2?.ToString(),
                //        ChangeDate = DateTime.Parse(cells.Cells[i, 10].Value2).Date,

                //        DeliveryAddress = cells.Cells[i, 16].Value2?.ToString(),
                //        OrderDate = cells.Cells[i, 17].Value2 == null ? null : DateTime.Parse(cells.Cells[i, 17].Value2).Date,

                //        WorkingEmail = cells.Cells[i, 29].Value2?.ToString(),
                //        PrivateEmail = cells.Cells[i, 30].Value2?.ToString(),

                //        WorkingPhone = cells.Cells[i, 32].Value2?.ToString(),
                //        MobilePhone = cells.Cells[i, 34].Value2?.ToString()
                //    };

                //    orders.Add(order);
            }

            return orders;
        }
    }
}
