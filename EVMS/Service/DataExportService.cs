using ClosedXML.Excel;
using System.Diagnostics;
using System.IO;

namespace EVMS.Service
{
    public class ParameterValue
    {
        public string Parameter { get; set; }
        public double Value { get; set; }
    }

    public class DataExportService
    {
        private readonly DataStorageService _dataService;
        private readonly string _baseExportPath;

        public DataExportService(DataStorageService dataService, string baseExportPath)
        {
            _dataService = dataService ?? throw new ArgumentNullException(nameof(dataService));
            _baseExportPath = baseExportPath ?? throw new ArgumentNullException(nameof(baseExportPath));
        }

        public string ExportDailyCumulativeReport(DateTime date, string partNo)
        {
            if (string.IsNullOrEmpty(partNo))
                throw new ArgumentException("Part number must be provided.", nameof(partNo));

            var partConfig = _dataService.GetPartConfigByPartNumber(partNo).ToList();

            if (partConfig == null || !partConfig.Any())
                throw new Exception($"No part config found for part number: {partNo}");

            foreach (var pc in partConfig)
            {
                Debug.WriteLine($"Parameter: {pc.Parameter}, Nominal: {pc.Nominal}, RTolPlus: {pc.RTolPlus}, RTolMinus: {pc.RTolMinus}");
            }

            var USL = partConfig.Select(p => new ParameterValue
            { Parameter = p.Parameter, Value = p.Nominal - p.RTolMinus }).ToList();

            var MEAN = partConfig.Select(p => new ParameterValue
            { Parameter = p.Parameter, Value = p.Nominal }).ToList();

            var LSL = partConfig.Select(p => new ParameterValue
            { Parameter = p.Parameter, Value = p.Nominal + p.RTolPlus }).ToList();

            string folder = Path.Combine(_baseExportPath, date.ToString("yyyy-MM-dd"));
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            string fileName = $"CumulativeReport_{partNo}_{date:yyyyMMdd}.xlsx";
            string filePath = Path.Combine(folder, fileName);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Report");

            // Set default font
            ws.Style.Font.FontName = "Segoe UI";
            ws.Style.Font.FontSize = 11;

            // Merge and style headers across columns A to E for modern centered look
            ws.Range("J1:M1").Merge();
            ws.Cell("J1").Value = "Company Name: SPRL";
            ws.Cell("J1").Style.Font.Bold = true;
            ws.Cell("J1").Style.Font.FontSize = 14;
            ws.Cell("J1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("J1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(1).Height = 25;

            ws.Range("J2:M2").Merge();
            ws.Cell("J2").Value = $"Date: {date:dd-MMM-yyyy}";
            ws.Cell("J2").Style.Font.Bold = true;
            ws.Cell("J2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("J2").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(2).Height = 20;

            ws.Range("J3:M3").Merge();
            ws.Cell("J3").Value = $"Part Number: {partNo}";
            ws.Cell("J3").Style.Font.Bold = true;
            ws.Cell("J3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("J3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(3).Height = 20;

            // Parameter header row with bold and light blue background
            // Parameter headers start at row 5, column 4 (D), move across
            for (int i = 0; i < partConfig.Count; i++)
            {
                var cell = ws.Cell(5, 4 + i);
                cell.Value = partConfig[i].Parameter;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D4E6F1");
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            // Labels (USL, MEAN, LSL) in column 1 (A) from rows 6,7,8
            var labelFormats = new (string Label, XLColor Color)[]
            {
    ("USL", XLColor.Red),
    ("MEAN", XLColor.ForestGreen),
    ("LSL", XLColor.Red)
            };

            for (int rowIdx = 0; rowIdx < labelFormats.Length; rowIdx++)
            {
                var labelCell = ws.Cell(6 + rowIdx, 3); // Column 1 (A)
                labelCell.Value = labelFormats[rowIdx].Label;
                labelCell.Style.Font.Bold = true;
                labelCell.Style.Font.FontColor = labelFormats[rowIdx].Color;
                labelCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            }

            // Fix width of label column for neatness
            ws.Column(1).Width = 15;

            // Fill data matching parameter header columns and label rows
            for (int i = 0; i < partConfig.Count; i++)
            {
                var uslCell = ws.Cell(6, 4 + i); // Row 6, parameter columns starting col 4 (D)
                uslCell.Value = USL[i].Value;
                uslCell.Style.Font.Bold = true;
                uslCell.Style.Font.FontColor = XLColor.Red;
                uslCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                uslCell.Style.NumberFormat.Format = "00.000;00.000;00.000";

                var meanCell = ws.Cell(7, 4 + i); // Row 7, col 4+
                meanCell.Value = MEAN[i].Value;
                meanCell.Style.Font.Bold = true;
                meanCell.Style.Font.FontColor = XLColor.ForestGreen;
                meanCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                meanCell.Style.NumberFormat.Format = "00.000;00.000;00.000";

                var lslCell = ws.Cell(8, 4 + i);  // Row 8 col 4+
                lslCell.Value = LSL[i].Value;
                lslCell.Style.Font.Bold = true;
                lslCell.Style.Font.FontColor = XLColor.Red;
                lslCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                lslCell.Style.NumberFormat.Format = "00.000;00.000;00.000";
            }

            // Calculate last column used by parameters
            int lastColumn = 4 + partConfig.Count - 1;

            // Select full data range: rows 5 to 8, columns 1 (A) to lastColumn (D + count -1)
            var dataRange = ws.Range(5, 3, 8, lastColumn);
            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Autofit all relevant columns (from column A to last parameter column)
            ws.Columns(1, lastColumn).AdjustToContents();

            // Optionally enforce minimum widths for data columns (starting from column 3 if you want)
            for (int col = 1; col <= lastColumn; col++)
            {
                if (ws.Column(col).Width < 8)
                    ws.Column(col).Width = 8;
            }




            // === Below: Add Measured Data after LSL row ===

            var yesterday = DateTime.Today.AddDays(-1);
            var measuredData = _dataService.GetAllMeasurementReadingsDynamic(partNo, yesterday);

            if (measuredData != null && measuredData.Count > 0)
            {
                int startRow = 11; // row after LSL

                // Write "SI No" header in column A
                ws.Cell(startRow, 1).Value = "SI No";
                ws.Cell(startRow, 1).Style.Font.Bold = true;
                ws.Cell(startRow, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                ws.Cell(startRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(startRow, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                var headers = measuredData[1].Keys.ToList();

                // Write other headers starting from column B
                for (int colIndex = 0; colIndex < headers.Count; colIndex++)
                {
                    var cell = ws.Cell(startRow, 2 + colIndex);
                    cell.Value = headers[colIndex];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // Write data rows with SI No
                for (int rowIndex = 0; rowIndex < measuredData.Count; rowIndex++)
                {
                    int excelRow = startRow + 1 + rowIndex;
                    ws.Cell(excelRow, 1).Value = rowIndex + 1;  // SI No
                    ws.Cell(excelRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(excelRow, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    var row = measuredData[rowIndex];
                    for (int colIndex = 0; colIndex < headers.Count; colIndex++)
                    {
                        var key = headers[colIndex];
                        var cell = ws.Cell(excelRow, 2 + colIndex);
                        var value = row[key];

                        if (value == null)
                        {
                            cell.Value = "";
                        }
                        else if (value is int || value is long)
                        {
                            cell.Value = Convert.ToInt64(value);
                            cell.Style.NumberFormat.Format = "0";  // Integer format
                        }
                        else if (value is float || value is double || value is decimal)
                        {
                            cell.Value = Convert.ToDouble(value);
                            cell.Style.NumberFormat.Format = "00.000;00.000;00.000";  // 3 decimal format
                        }
                        else if (value is DateTime dt)
                        {
                            cell.Value = dt;
                            cell.Style.DateFormat.Format = "dd-MMM-yyyy HH:mm:ss";
                        }
                        else
                        {
                            // Default as string
                            cell.Value = value.ToString();
                        }

                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }
                }

                // Autofit all columns including SI No
                ws.Columns(1, 1 + headers.Count).AdjustToContents();

                // Enforce minimum column width (e.g., 10) to avoid ##### display
                int minWidth = 20;
                int maxCol = 1 + headers.Count;
                for (int col = 1; col <= maxCol; col++)
                {
                    if (ws.Column(col).Width < minWidth)
                        ws.Column(col).Width = minWidth;
                }

                // Save the workbook
                wb.SaveAs(filePath);


            }
            return filePath;
        }
    }
}
