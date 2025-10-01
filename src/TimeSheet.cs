using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Tempo.Exporter
{
    internal class TimeSheet
    {
        private string _path;
        private static readonly TimeSpan MinimumBreak = TimeSpan.FromMinutes(30);

        public TimeSheet(string path)
        {
            _path = path;
        }

        public void ExportWorkingTime(Dictionary<DateTime, IEnumerable<TimeRange>> workingDays)
        {
            var excelWorkbook = ReadExcelFile(_path);
            var excelSheet = GetLastVisibleSheet(excelWorkbook);

            if (excelSheet is null)
            {
                Console.WriteLine("No visible sheet has been found in the workbook. Time export has been aborted.");
                return;
            }

            foreach (var currentWorkingDay in workingDays)
            {
                var currentDateCell = FindDateCell(excelSheet, currentWorkingDay.Key);

                if (currentDateCell is null)
                {
                    continue;
                }

                var beginCell = GetCell(excelSheet, currentDateCell.RowIndex, currentDateCell.ColumnIndex + ArbeitszeitenConstants.OFFSET_BEGIN_CELL);
                var endCell = GetCell(excelSheet, currentDateCell.RowIndex, currentDateCell.ColumnIndex + ArbeitszeitenConstants.OFFSET_END_CELL);
                var pauseCell = GetCell(excelSheet, currentDateCell.RowIndex, currentDateCell.ColumnIndex + ArbeitszeitenConstants.OFFSET_PAUSE_CELL);

                var workBegin = TimeRange.GetWorkBegin(currentWorkingDay.Value);
                var workEnd = TimeRange.GetWorkEnd(currentWorkingDay.Value);
                var pauseMinutes = TimeRange.CalculateBreaks(currentWorkingDay.Value);

                if (currentWorkingDay.Value.All(wd => wd.IssueId == ArbeitszeitenConstants.KRANK_ISSUE_ID))
                {
                    beginCell.SetCellType(CellType.Blank);
                    endCell.SetCellType(CellType.Blank);
                    SetBemerkung(excelSheet, currentDateCell, "Krank");
                    continue;
                }

                if (currentWorkingDay.Value.All(wd => wd.IssueId == ArbeitszeitenConstants.URLAUB_ISSUE_ID))
                {
                    beginCell.SetCellType(CellType.Blank);
                    endCell.SetCellType(CellType.Blank);
                    SetBemerkung(excelSheet, currentDateCell, "Urlaub");
                    continue;
                }

                // Append or extend break if not recorded
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

            XSSFFormulaEvaluator.EvaluateAllFormulaCells(excelWorkbook);

            SaveExcelFile(_path, excelWorkbook);
        }

        private static ISheet GetLastVisibleSheet(IWorkbook workbook)
        {
            ISheet sheet = null;

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                if (workbook.GetSheetVisibility(i) != SheetVisibility.Hidden)
                {
                    sheet = workbook.GetSheetAt(i);
                }
            }

            return sheet;
        }

        private static ICell GetCell(ISheet sheet, int rowIndex, int columnIndex)
        {
            var cell = sheet.GetRow(rowIndex).GetCell(columnIndex);

            if (cell is null)
            {
                return sheet.GetRow(rowIndex).CreateCell(columnIndex);
            }

            return cell;
        }

        private static void SetBemerkung(ISheet sheet, ICell currentDateCell, string text)
        {
            var bemerkungCell = GetCell(sheet, currentDateCell.RowIndex, currentDateCell.ColumnIndex + ArbeitszeitenConstants.OFFSET_BEMERKUNG_CELL);
            bemerkungCell.SetCellValue(text);
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
