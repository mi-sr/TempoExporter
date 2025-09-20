using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Tempo.Exporter
{
    internal class TimeSheet(string path)
    {
        private readonly string _path = path;
        private static readonly TimeSpan MinimumBreak = TimeSpan.FromMinutes(30);

        public void ExportWorkingTime(Dictionary<DateTime, IEnumerable<TimeRange>> workingDays)
        {
            var excelWorkbook = ReadExcelFile(_path);
            var excelSheet = excelWorkbook.GetSheetAt(excelWorkbook.NumberOfSheets - 1);

            foreach (var currentWorkingDay in workingDays)
            {
                var currentDateCell = FindDateCell(excelSheet, currentWorkingDay.Key);

                if (currentDateCell is null)
                {
                    continue;
                }

                var beginCell = excelSheet.GetRow(currentDateCell.RowIndex).GetCell(currentDateCell.ColumnIndex + 2);
                var endCell = excelSheet.GetRow(currentDateCell.RowIndex).GetCell(currentDateCell.ColumnIndex + 3);
                var pauseCell = excelSheet.GetRow(currentDateCell.RowIndex).GetCell(currentDateCell.ColumnIndex + 5);

                var workBegin = TimeRange.GetWorkBegin(currentWorkingDay.Value);
                var workEnd = TimeRange.GetWorkEnd(currentWorkingDay.Value);
                var pauseMinutes = TimeRange.CalculateBreaks(currentWorkingDay.Value);

                // Append break if not recorded
                if (workEnd - workBegin >= TimeSpan.FromHours(6) && TimeSpan.FromMinutes(pauseMinutes) < MinimumBreak)
                {
                    var difference = MinimumBreak - TimeSpan.FromMinutes(pauseMinutes);
                    
                    pauseMinutes = (int)MinimumBreak.TotalMinutes;
                    workEnd = workEnd.Add(difference);
                }

                SetTimeValue(excelWorkbook, workBegin, beginCell);
                SetTimeValue(excelWorkbook, workEnd, endCell);

                if (pauseMinutes > 0)
                {
                    // Override formula with custom value
                    pauseCell.SetCellFormula(null);
                    pauseCell.SetCellValue(pauseMinutes);
                }
            }

            BaseFormulaEvaluator.EvaluateAllFormulaCells(excelWorkbook);

            SaveExcelFile(_path, excelWorkbook);
        }

        private static void SetTimeValue(IWorkbook workbook, TimeSpan? value, ICell cell)
        {
            if (value is null)
            {
                return;
            }

            cell.SetCellValue(value.Value.TotalDays);
            cell.CellStyle = GetTimespanCellFormat(workbook);
        }

        private static ICellStyle GetTimespanCellFormat(IWorkbook workbook)
        {
            var dataformat = workbook.CreateDataFormat();
            ICellStyle style = workbook.CreateCellStyle();
            style.DataFormat = dataformat.GetFormat("hh:mm");

            return style;
        }

        private IWorkbook ReadExcelFile(string path)
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            return new XSSFWorkbook(fs);
        }

        private void SaveExcelFile(string path, IWorkbook workbook)
        {
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            workbook.Write(fs);
        }

        private static ICell FindDateCell(ISheet sheet, DateTime date)
        {
            for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);

                if (row == null)
                {
                    continue;
                }

                for (int cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; cellIndex++)
                {
                    ICell cell = row.GetCell(cellIndex);

                    if (cell == null)
                    {
                        continue;
                    }

                    DateTime? cellValue = null;

                    try
                    {
                        cellValue = cell.DateCellValue;
                    }
                    catch
                    {
                        // ignore
                    }

                    if (cellValue is null)
                    {
                        continue;
                    }

                    if (cellValue == date)
                    {
                        return cell;
                    }
                }
            }

            return null;
        }
    }
}
