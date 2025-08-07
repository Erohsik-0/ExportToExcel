using ExportToExcel.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ClosedXML.Excel;
using System.Diagnostics;

namespace ExportToExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index() => View();

        public IActionResult Privacy() => View();

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpGet("ExportToExcel")]
        public IActionResult ExportToExcel()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "jsondata.json");

            if (!System.IO.File.Exists(filePath))
                return NotFound("JSON data not found");

            var jsonString = System.IO.File.ReadAllText(filePath);

            var dataList = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonString);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Data");

                if (dataList != null && dataList.Count > 0)
                {
                    // Preserve original order of headers from first item
                    var headers = dataList[0].Keys.ToList();

                    // Add headers
                    for (int col = 0; col < headers.Count; col++)
                    {
                        var headerCell = worksheet.Cell(1, col + 1);
                        headerCell.Value = headers[col];
                        StyleHeaderCell(headerCell);
                    }

                    worksheet.Range(1, 1, 1, headers.Count).SetAutoFilter();
                    worksheet.SheetView.FreezeRows(1);

                    // Add data
                    for (int row = 0; row < dataList.Count; row++)
                    {
                        var rowData = dataList[row];
                        for (int col = 0; col < headers.Count; col++)
                        {
                            var key = headers[col];
                            if (rowData.TryGetValue(key, out var value))
                            {
                                var cell = worksheet.Cell(row + 2, col + 1);
                                SetCellValueWithType(cell, value);
                                StyleDataCell(cell);
                            }
                        }
                    }

                    worksheet.Columns().AdjustToContents();
                }

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                var content = stream.ToArray();

                string filename = $"Export_{DateTime.Now:yyyy-MM-dd_HHmm}.xlsx";
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
            }
        }

        [HttpGet("ExportToExcel2")]
        public IActionResult ExportToExcel2()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "cosmosdata.json");

            if (!System.IO.File.Exists(filePath))
                return NotFound("JSON data not found");

            var jsonString = System.IO.File.ReadAllText(filePath);

            List<Dictionary<string, object>> flatList = new();
            List<string> orderedHeaders = new(); // Track header order

            Dictionary<string, object> FlattenJson(JToken token, string parentPath = "", bool isFirstItem = false)
            {
                var result = new Dictionary<string, object>();

                if (token is JObject jObj)
                {
                    foreach (var prop in jObj.Properties())
                    {
                        string path = string.IsNullOrEmpty(parentPath) ? prop.Name : $"{parentPath}.{prop.Name}";
                        var value = prop.Value;

                        if (value.Type == JTokenType.Object)
                        {
                            foreach (var nested in FlattenJson(value, path, isFirstItem))
                            {
                                result[nested.Key] = nested.Value;
                                if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                    orderedHeaders.Add(nested.Key);
                            }
                        }
                        else if (value.Type == JTokenType.Array)
                        {
                            var array = value as JArray;
                            for (int i = 0; i < array.Count; i++)
                            {
                                if (array[i] is JObject itemObj)
                                {
                                    foreach (var nested in FlattenJson(itemObj, $"{path}[{i}]", isFirstItem))
                                    {
                                        result[nested.Key] = nested.Value;
                                        if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                            orderedHeaders.Add(nested.Key);
                                    }
                                }
                                else
                                {
                                    string arrayKey = $"{path}[{i}]";
                                    result[arrayKey] = ((JValue)array[i]).Value;
                                    if (isFirstItem && !orderedHeaders.Contains(arrayKey))
                                        orderedHeaders.Add(arrayKey);
                                }
                            }
                        }
                        else
                        {
                            result[path] = ((JValue)value).Value;
                            if (isFirstItem && !orderedHeaders.Contains(path))
                                orderedHeaders.Add(path);
                        }
                    }
                }

                return result;
            }

            try
            {
                var token = JToken.Parse(jsonString);
                bool isFirstItem = true;

                if (token is JObject obj)
                {
                    if (obj["Documents"] is JArray docs)
                    {
                        foreach (var item in docs)
                        {
                            flatList.Add(FlattenJson(item, "", isFirstItem));
                            isFirstItem = false;
                        }
                    }
                    else if (obj["Items"] is JArray items)
                    {
                        foreach (var item in items)
                        {
                            flatList.Add(FlattenJson(item, "", isFirstItem));
                            isFirstItem = false;
                        }
                    }
                    else
                    {
                        flatList.Add(FlattenJson(obj, "", isFirstItem));
                    }
                }
                else if (token is JArray arr)
                {
                    foreach (var item in arr)
                    {
                        flatList.Add(FlattenJson(item, "", isFirstItem));
                        isFirstItem = false;
                    }
                }
            }
            catch (JsonReaderException ex)
            {
                return BadRequest($"Invalid JSON format: {ex.Message}");
            }

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Data");

            if (flatList.Count > 0)
            {
                // Use ordered headers, then add any additional headers found in other items
                var allHeaders = orderedHeaders.ToList();
                var additionalHeaders = flatList.SelectMany(d => d.Keys)
                                                .Distinct()
                                                .Where(h => !allHeaders.Contains(h))
                                                .OrderBy(h => h)
                                                .ToList();
                allHeaders.AddRange(additionalHeaders);

                // Add headers
                for (int col = 0; col < allHeaders.Count; col++)
                {
                    var headerCell = worksheet.Cell(1, col + 1);
                    headerCell.Value = allHeaders[col];
                    StyleHeaderCell(headerCell);
                }

                worksheet.Range(1, 1, 1, allHeaders.Count).SetAutoFilter();
                worksheet.SheetView.FreezeRows(1);

                // Add data rows
                for (int row = 0; row < flatList.Count; row++)
                {
                    var rowData = flatList[row];
                    for (int col = 0; col < allHeaders.Count; col++)
                    {
                        rowData.TryGetValue(allHeaders[col], out var value);
                        var cell = worksheet.Cell(row + 2, col + 1);
                        SetCellValueWithType(cell, value);
                        StyleDataCell(cell);
                    }
                }

                worksheet.Columns().AdjustToContents();
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            string filename = $"CosmosExport_{DateTime.Now:yyyy-MM-dd_HHmm}.xlsx";
            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        }

        [HttpGet("ExportToExcelGrouped")]
        public IActionResult ExportToExcelGrouped(string groupByField = "category")
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "jsondata.json");

            if (!System.IO.File.Exists(filePath))
                return NotFound("JSON data not found");

            var jsonString = System.IO.File.ReadAllText(filePath);
            var dataList = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonString);

            if (dataList == null || dataList.Count == 0)
                return BadRequest("No data found in JSON file");

            using var workbook = new XLWorkbook();

            // Preserve original header order from first item
            var headers = dataList[0].Keys.ToList();

            // Group data by the specified field
            var groupedData = dataList
                .Where(item => item.ContainsKey(groupByField))
                .GroupBy(item => item[groupByField]?.ToString() ?? "Unknown")
                .OrderBy(g => g.Key)
                .ToList();

            // Handle items without the grouping field
            var ungroupedData = dataList.Where(item => !item.ContainsKey(groupByField)).ToList();

            // Create sheets for each group
            foreach (var group in groupedData)
            {
                CreateWorksheet(workbook, SanitizeSheetName(group.Key), group.ToList(), headers);
            }

            // Create sheet for ungrouped data if any exists
            if (ungroupedData.Count > 0)
            {
                CreateWorksheet(workbook, "Ungrouped", ungroupedData, headers);
            }

            // Create summary sheet
            CreateSummarySheet(workbook, groupedData, ungroupedData.Count, groupByField);

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            string filename = $"GroupedExport_{groupByField}_{DateTime.Now:yyyy-MM-dd_HHmm}.xlsx";
            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        }

        [HttpGet("ExportToExcelGroupedCosmos")]
        public IActionResult ExportToExcelGroupedCosmos(string groupByField = "type")
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "cosmosdata.json");

            if (!System.IO.File.Exists(filePath))
                return NotFound("JSON data not found");

            var jsonString = System.IO.File.ReadAllText(filePath);
            List<Dictionary<string, object>> flatList = new();
            List<string> orderedHeaders = new();

            Dictionary<string, object> FlattenJson(JToken token, string parentPath = "", bool isFirstItem = false)
            {
                var result = new Dictionary<string, object>();

                if (token is JObject jObj)
                {
                    foreach (var prop in jObj.Properties())
                    {
                        string path = string.IsNullOrEmpty(parentPath) ? prop.Name : $"{parentPath}.{prop.Name}";
                        var value = prop.Value;

                        if (value.Type == JTokenType.Object)
                        {
                            foreach (var nested in FlattenJson(value, path, isFirstItem))
                            {
                                result[nested.Key] = nested.Value;
                                if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                    orderedHeaders.Add(nested.Key);
                            }
                        }
                        else if (value.Type == JTokenType.Array)
                        {
                            var array = value as JArray;
                            for (int i = 0; i < array.Count; i++)
                            {
                                if (array[i] is JObject itemObj)
                                {
                                    foreach (var nested in FlattenJson(itemObj, $"{path}[{i}]", isFirstItem))
                                    {
                                        result[nested.Key] = nested.Value;
                                        if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                            orderedHeaders.Add(nested.Key);
                                    }
                                }
                                else
                                {
                                    string arrayKey = $"{path}[{i}]";
                                    result[arrayKey] = ((JValue)array[i]).Value;
                                    if (isFirstItem && !orderedHeaders.Contains(arrayKey))
                                        orderedHeaders.Add(arrayKey);
                                }
                            }
                        }
                        else
                        {
                            result[path] = ((JValue)value).Value;
                            if (isFirstItem && !orderedHeaders.Contains(path))
                                orderedHeaders.Add(path);
                        }
                    }
                }

                return result;
            }

            try
            {
                var token = JToken.Parse(jsonString);
                bool isFirstItem = true;

                if (token is JObject obj)
                {
                    if (obj["Documents"] is JArray docs)
                    {
                        foreach (var item in docs)
                        {
                            flatList.Add(FlattenJson(item, "", isFirstItem));
                            isFirstItem = false;
                        }
                    }
                    else if (obj["Items"] is JArray items)
                    {
                        foreach (var item in items)
                        {
                            flatList.Add(FlattenJson(item, "", isFirstItem));
                            isFirstItem = false;
                        }
                    }
                    else
                    {
                        flatList.Add(FlattenJson(obj, "", isFirstItem));
                    }
                }
                else if (token is JArray arr)
                {
                    foreach (var item in arr)
                    {
                        flatList.Add(FlattenJson(item, "", isFirstItem));
                        isFirstItem = false;
                    }
                }
            }
            catch (JsonReaderException ex)
            {
                return BadRequest($"Invalid JSON format: {ex.Message}");
            }

            if (flatList.Count == 0)
                return BadRequest("No data found after flattening");

            using var workbook = new XLWorkbook();

            // Use ordered headers, then add any additional headers found in other items
            var allHeaders = orderedHeaders.ToList();
            var additionalHeaders = flatList.SelectMany(d => d.Keys)
                                            .Distinct()
                                            .Where(h => !allHeaders.Contains(h))
                                            .OrderBy(h => h)
                                            .ToList();
            allHeaders.AddRange(additionalHeaders);

            // Group data by the specified field
            var groupedData = flatList
                .Where(item => item.ContainsKey(groupByField))
                .GroupBy(item => item[groupByField]?.ToString() ?? "Unknown")
                .OrderBy(g => g.Key)
                .ToList();

            // Handle items without the grouping field
            var ungroupedData = flatList.Where(item => !item.ContainsKey(groupByField)).ToList();

            // Create sheets for each group
            foreach (var group in groupedData)
            {
                CreateWorksheet(workbook, SanitizeSheetName(group.Key), group.ToList(), allHeaders);
            }

            // Create sheet for ungrouped data if any exists
            if (ungroupedData.Count > 0)
            {
                CreateWorksheet(workbook, "Ungrouped", ungroupedData, allHeaders);
            }

            // Create summary sheet
            CreateSummarySheet(workbook, groupedData, ungroupedData.Count, groupByField);

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            string filename = $"CosmosGroupedExport_{groupByField}_{DateTime.Now:yyyy-MM-dd_HHmm}.xlsx";
            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        }

        private void CreateWorksheet(XLWorkbook workbook, string sheetName, List<Dictionary<string, object>> data, List<string> headers)
        {
            var worksheet = workbook.Worksheets.Add(sheetName);

            if (data.Count == 0)
            {
                worksheet.Cell(1, 1).Value = "No data available";
                return;
            }

            // Add headers
            for (int col = 0; col < headers.Count; col++)
            {
                var headerCell = worksheet.Cell(1, col + 1);
                headerCell.Value = headers[col];
                StyleHeaderCell(headerCell);
            }

            worksheet.Range(1, 1, 1, headers.Count).SetAutoFilter();
            worksheet.SheetView.FreezeRows(1);

            // Add data rows
            for (int row = 0; row < data.Count; row++)
            {
                var rowData = data[row];
                for (int col = 0; col < headers.Count; col++)
                {
                    var key = headers[col];
                    if (rowData.TryGetValue(key, out var value))
                    {
                        var cell = worksheet.Cell(row + 2, col + 1);
                        SetCellValueWithType(cell, value);
                        StyleDataCell(cell);
                    }
                    else
                    {
                        var cell = worksheet.Cell(row + 2, col + 1);
                        cell.Value = "";
                        StyleDataCell(cell);
                    }
                }
            }

            worksheet.Columns().AdjustToContents();
        }

        private void CreateSummarySheet(XLWorkbook workbook, List<IGrouping<string, Dictionary<string, object>>> groupedData, int ungroupedCount, string groupByField)
        {
            var summarySheet = workbook.Worksheets.Add("Summary");

            // Title
            var titleCell = summarySheet.Cell(1, 1);
            titleCell.Value = $"Export Summary - Grouped by {groupByField}";
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 14;
            titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            summarySheet.Range(1, 1, 1, 3).Merge();

            // Headers
            summarySheet.Cell(3, 1).Value = "Group";
            summarySheet.Cell(3, 2).Value = "Record Count";
            summarySheet.Cell(3, 3).Value = "Sheet Name";

            for (int col = 1; col <= 3; col++)
            {
                StyleHeaderCell(summarySheet.Cell(3, col));
            }

            int row = 4;
            int totalRecords = 0;

            // Group data
            foreach (var group in groupedData)
            {
                summarySheet.Cell(row, 1).Value = group.Key;
                summarySheet.Cell(row, 2).Value = group.Count();
                summarySheet.Cell(row, 3).Value = SanitizeSheetName(group.Key);

                for (int col = 1; col <= 3; col++)
                {
                    StyleDataCell(summarySheet.Cell(row, col));
                }

                totalRecords += group.Count();
                row++;
            }

            // Ungrouped data
            if (ungroupedCount > 0)
            {
                summarySheet.Cell(row, 1).Value = "Ungrouped";
                summarySheet.Cell(row, 2).Value = ungroupedCount;
                summarySheet.Cell(row, 3).Value = "Ungrouped";

                for (int col = 1; col <= 3; col++)
                {
                    StyleDataCell(summarySheet.Cell(row, col));
                }

                totalRecords += ungroupedCount;
                row++;
            }

            // Total
            summarySheet.Cell(row + 1, 1).Value = "Total Records:";
            summarySheet.Cell(row + 1, 2).Value = totalRecords;
            summarySheet.Cell(row + 1, 1).Style.Font.Bold = true;
            summarySheet.Cell(row + 1, 2).Style.Font.Bold = true;
            StyleDataCell(summarySheet.Cell(row + 1, 1));
            StyleDataCell(summarySheet.Cell(row + 1, 2));

            summarySheet.Columns().AdjustToContents();
        }

        private string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Unknown";

            // Remove invalid characters for Excel sheet names
            var invalidChars = new char[] { '/', '\\', '?', '*', '[', ']', ':' };
            foreach (var invalidChar in invalidChars)
            {
                name = name.Replace(invalidChar, '_');
            }

            // Excel sheet names must be 31 characters or less
            if (name.Length > 31)
                name = name.Substring(0, 31);

            return name;
        }

        private void StyleHeaderCell(IXLCell cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        private void StyleDataCell(IXLCell cell)
        {
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        private void SetCellValueWithType(IXLCell cell, object value)
        {
            if (value == null)
            {
                cell.Value = "";
                return;
            }

            if (value is bool b)
            {
                cell.Value = b;
            }
            else if (value is int i)
            {
                cell.Value = i;
            }
            else if (value is long l)
            {
                cell.Value = l;
            }
            else if (value is double d)
            {
                cell.Value = d;
            }
            else if (value is decimal m)
            {
                cell.Value = m;
            }
            else if (value is DateTime dt)
            {
                cell.Value = dt;
                cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
            }
            else
            {
                var str = value.ToString();

                // Try parse for string-based values
                if (bool.TryParse(str, out var boolParsed))
                    cell.Value = boolParsed;
                else if (int.TryParse(str, out var intParsed))
                    cell.Value = intParsed;
                else if (long.TryParse(str, out var longParsed))
                    cell.Value = longParsed;
                else if (decimal.TryParse(str, out var decimalParsed))
                    cell.Value = decimalParsed;
                else if (DateTime.TryParse(str, out var dtParsed))
                {
                    cell.Value = dtParsed;
                    cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
                }
                else
                {
                    cell.Value = str;
                }
            }
        }
    }
}