namespace DeflectorCheck
{
    using System.Runtime.InteropServices;
    using System.Runtime.Versioning;
    using Excel = Microsoft.Office.Interop.Excel;

    internal class Spreadsheet
    {
        private const double PersonalThreshold = 0.75;
        private const double CoopThreshold = 0.75;
        private const double ContractThreshold = 0.75;
        private const int SmallContractSizeThreshold = 100;

        private static readonly string RoamingFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DeflectorCheck\";

        private readonly List<Excel.Worksheet> worksheets = new ();

        public void Create()
        {
            Console.WriteLine("Creating Spreadsheet...");
            Logger.Info("Creating Spreadsheet...");

            Excel.Application xlApp = new ();
            Excel.Workbook workbook = xlApp.Workbooks.Add();

            Excel.Worksheet summary;
            summary = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            summary.Name = "Summary";
            worksheets.Add(summary);

            Excel.Worksheet slotted = AddSheet(workbook, "Slotted", summary);
            Excel.Worksheet ratio = AddSheet(workbook, "Personal Ratio", slotted);
            Excel.Worksheet coop = AddSheet(workbook, "Coop Ratio", ratio);
            Excel.Worksheet contract = AddSheet(workbook, "Contract Ratio", coop);

            List<string> contractNames = ContractDetail.GetSortedContracts();

            int headerRow = 2;
            int headerCol = 2;

            int dataStartRow;
            int dataEndRow;
            int dataStartCol;
            int dataEndCol;

            int row = headerRow;
            int col = headerCol;

            SetCellOnAllSheets(row, col, "Contract", header: true);
            row++;
            SetCellOnAllSheets(row, col, "Coop Size", header: true);
            row++;
            SetCellOnAllSheets(row, col, "Token Timer", header: true);
            row++;

            dataStartRow = row;

            foreach (string member in GroupData.GetSortedMembers())
            {
                SetCellOnAllSheets(row, col, GroupData.MemberEggName(member));
                row++;
            }

            row = headerRow;
            dataStartCol = col + 1;

            foreach (string contractName in ContractDetail.GetSortedContracts())
            {
                Logger.Info("Adding " + contractName + " to spreadsheet");

                row = headerRow;
                col++;

                SetCellOnAllSheets(row, col, contractName, header: true);
                row++;
                SetCellOnAllSheets(row, col, ContractDetail.GetContractSize(contractName).ToString());
                row++;
                SetCellOnAllSheets(row, col, ContractDetail.GetContractTokenTimer(contractName).ToString());

                // We will increment row in the foreach
                foreach (string member in GroupData.GetSortedMembers())
                {
                    row++;

                    if (GroupData.MemberHasContract(member, contractName))
                    {
                        Excel.Range formatRange;

                        string shit = "IF(OR(Slotted!" + GetColumnLetter(col) + row + " = \"Yes\", " +
                                      "'Personal Ratio'!" + GetColumnLetter(col) + row + " > Personal, " +
                                      "'Coop Ratio'!" + GetColumnLetter(col) + row + " > Coop, " +
                                      "'Contract Ratio'!" + GetColumnLetter(col) + row + " > Contract), \"Yes\", \"No\")";
                        summary.Cells[row, col] = "=IF(" + GetColumnLetter(col) + 3 + "<=Small," + shit + ",\"\")";

                        if (GroupData.MemberDeflecting(member, contractName))
                        {
                            slotted.Cells[row, col] = "=IF(" + GetColumnLetter(col) + 3 + "<=Small,\"Yes\",\"\")";
                        }
                        else
                        {
                            slotted.Cells[row, col] = "=IF(" + GetColumnLetter(col) + 3 + "<=Small,\"No\",\"\")";
                        }

                        ratio.Cells[row, col] = "=IF(" + GetColumnLetter(col) + 3 + "<=Small," + GroupData.MemberPersonalRatio(member, contractName) + ",\"\")";
                        formatRange = ratio.Cells[row, col];
                        formatRange.NumberFormat = "0.00%";

                        coop.Cells[row, col] = "=IF(" + GetColumnLetter(col) + 3 + "<=Small," + GroupData.MemberCoopRatio(member, contractName) + ",\"\")";
                        formatRange = coop.Cells[row, col];
                        formatRange.NumberFormat = "0.00%";

                        contract.Cells[row, col] = "=IF(" + GetColumnLetter(col) + 3 + "<=Small," + GroupData.MemberContractRatio(member, contractName) + ",\"\")";
                        formatRange = contract.Cells[row, col];
                        formatRange.NumberFormat = "0.00%";
                    }
                }
            }

            dataEndRow = row;
            dataEndCol = col;

            col++;
            SetCellOnAllSheets(headerRow, col, "Average", header: true);
            MergeCellsOnAllSheets(headerRow, col, headerRow + 2, col);

            row = headerRow + 2;
            while (row < dataEndRow)
            {
                row++;

                string range = GetRangeForFormula(dataStartCol, dataEndCol, row);

                Excel.Range formulaRange;

                formulaRange = summary.Cells[row, col];
                formulaRange.Formula = "=IFERROR(COUNTIF(" + range + ", \"Yes\") / ( COUNTIF(" + range + ", \"Yes\") + COUNTIF(" + range + ", \"No\")),\"\")";
                formulaRange.NumberFormat = "0.00%";

                formulaRange = slotted.Cells[row, col];

                formulaRange.Formula = "=IFERROR(COUNTIF(" + range + ", \"Yes\") / ( COUNTIF(" + range + ", \"Yes\") + COUNTIF(" + range + ", \"No\")),\"\")";
                formulaRange.NumberFormat = "0.00%";

                formulaRange = ratio.Cells[row, col];
                formulaRange.Formula = "=IFERROR(AVERAGE(" + range + "),\"\")";
                formulaRange.NumberFormat = "0.00%";

                formulaRange = coop.Cells[row, col];
                formulaRange.Formula = "=IFERROR(AVERAGE(" + range + "),\"\")";
                formulaRange.NumberFormat = "0.00%";

                formulaRange = contract.Cells[row, col];
                formulaRange.Formula = "=IFERROR(AVERAGE(" + range + "),\"\")";
                formulaRange.NumberFormat = "0.00%";
            }

            SetColorScaleConditional(new List<Excel.Worksheet>() { ratio, coop, contract }, dataStartRow, dataStartCol, dataEndRow, dataEndCol);
            SetTextConditional(new List<Excel.Worksheet>() { summary, slotted }, "No", dataStartRow, dataStartCol, dataEndRow, dataEndCol);

            foreach (Excel.Worksheet sheet in worksheets)
            {
                int borderCol = col;
                int borderDataCol = dataEndCol;
                int headerEndCol = borderCol - 1;

                // Border for all cells with content
                SetBorder(sheet, headerRow, headerCol, row, borderCol);

                // Border around all content
                SetOutsideBorder(sheet, headerRow, headerCol, row, borderCol);

                // Border around the data
                SetOutsideBorder(sheet, dataStartRow, dataStartCol, dataEndRow, borderDataCol);

                // Border around header
                SetOutsideBorder(sheet, headerRow, headerCol, dataStartRow - 1, borderCol);

                // Border around the contract specific header
                SetOutsideBorder(sheet, headerRow, headerCol + 1, dataStartRow - 1, headerEndCol);

                Excel.Range filterRange = sheet.Rows[2];
                filterRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

                sheet.Activate();
                sheet.Application.ActiveWindow.SplitRow = 4;
                sheet.Application.ActiveWindow.SplitColumn = 2;
                sheet.Application.ActiveWindow.FreezePanes = true;
            }

            Excel.Worksheet key;
            key = workbook.Worksheets.Add(Before: summary);
            key.Name = "Settings";
            worksheets.Add(key);

            key.Cells[2, 2] = "Personal Ratio Pass Threshold";
            key.Cells[3, 2] = "Coop Ratio Pass Threshold";
            key.Cells[4, 2] = "Contract Ratio Pass Threshold";
            key.Cells[5, 2] = "Max Contract Size Threshold";

            key.Cells[2, 3] = PersonalThreshold;
            key.Cells[3, 3] = CoopThreshold;
            key.Cells[4, 3] = ContractThreshold;
            key.Cells[5, 3] = SmallContractSizeThreshold;

            Excel.Range nameRange = key.Cells[2, 3];
            nameRange.Name = "Personal";
            nameRange.NumberFormat = "0.00%";
            nameRange = key.Cells[3, 3];
            nameRange.Name = "Coop";
            nameRange.NumberFormat = "0.00%";
            nameRange = key.Cells[4, 3];
            nameRange.Name = "Contract";
            nameRange.NumberFormat = "0.00%";
            nameRange = key.Cells[5, 3];
            nameRange.Name = "Small";

            SetBorder(key, 2, 2, 5, 3);
            SetOutsideBorder(key, 2, 2, 5, 3);

            AutoFixAllSheets();

            summary.Activate();

            Console.WriteLine("Done creating Spreadsheet...");
            Logger.Info("Done creating Spreadsheet...");

            try
            {
                string filename = RoamingFolder + Config.GroupName + ".xlsx";
                Logger.Info("Saving Spreadsheet as " + filename);
                workbook.SaveAs(
                    filename,
                    Excel.XlFileFormat.xlWorkbookDefault,
                    ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
            }
            catch (Exception e)
            {
                // Should we do something
                Console.WriteLine("Error on savings spreadsheet");
                Console.WriteLine(e.ToString());
                Logger.Error("Error on saving spreadsheet", e);
            }

            try
            {
                workbook.Close(SaveChanges: false);
            }
            catch (Exception e)
            {
                // Should we do something
                Console.WriteLine("Error on closing spreadsheet");
                Console.WriteLine(e.ToString());
                Logger.Error("Error on closing spreadsheet", e);
            }

            xlApp.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            /*
            if (Marshal.FinalReleaseComObject(summary) != 0)
            {
                Console.WriteLine("Error on releasing summary");
            }

            if (Marshal.FinalReleaseComObject(slotted) != 0)
            {
                Console.WriteLine("Error on releasing slotted");
            }

            if (Marshal.FinalReleaseComObject(ratio) != 0)
            {
                Console.WriteLine("Error on releasing ratio");
            }

            if (Marshal.FinalReleaseComObject(coop) != 0)
            {
                Console.WriteLine("Error on releasing coop");
            }

            if (Marshal.FinalReleaseComObject(contract) != 0)
            {
                Console.WriteLine("Error on releasing contract");
            }

            if (Marshal.FinalReleaseComObject(key) != 0)
            {
                Console.WriteLine("Error on releasing key");
            }

            if (Marshal.FinalReleaseComObject(workbook) != 0)
            {
                Console.WriteLine("Error on releasing workbook");
            }

            if (Marshal.FinalReleaseComObject(xlApp) != 0)
            {
                Console.WriteLine("Error on releasing app");
            }
            */

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void SetCellHighlight(Excel.Worksheet sheet, int row, int col, System.Drawing.KnownColor color)
        {
            Excel.Range formatRange = sheet.Cells[row, col];
            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromKnownColor(color));
        }

        private static void SetCellAsHeader(Excel.Worksheet sheet, int row, int col)
        {
            Excel.Range formatRange = sheet.Cells[row, col];
            formatRange.Font.Bold = true;
        }

        private static void SetColorScaleConditional(List<Excel.Worksheet> sheets, int startRow, int startCol, int stopRow, int stopCol)
        {
            foreach (Excel.Worksheet sheet in sheets)
            {
                Excel.Range formatRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[stopRow, stopCol]];
                Excel.ColorScale colorScale = formatRange.FormatConditions.AddColorScale(3);
                colorScale.ColorScaleCriteria[1].FormatColor.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                colorScale.ColorScaleCriteria[2].FormatColor.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                colorScale.ColorScaleCriteria[3].FormatColor.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            }
        }

        private static void SetTextConditional(List<Excel.Worksheet> sheets, string condition, int startRow, int startCol, int stopRow, int stopCol)
        {
            foreach (Excel.Worksheet sheet in sheets)
            {
                Excel.Range formatRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[stopRow, stopCol]];
                Excel.FormatCondition formatCondition = formatRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, condition);
                formatCondition.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }
        }

        private static void MergeCells(Excel.Worksheet sheet, int startRow, int startCol, int stopRow, int stopCol)
        {
            Excel.Range formatRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[stopRow, stopCol]];
            formatRange.Merge();
        }

        private static string GetColumnLetter(int col)
        {
            string result = string.Empty;

            int tempCol = col - 1;

            while (tempCol / 26 > 0)
            {
                char c1 = (char)((tempCol / 26) - 1 + (int)'A');
                result += c1.ToString();
                tempCol %= 26;
            }

            char c2 = (char)(tempCol + (int)'A');
            result += c2.ToString();

            return result;
        }

        private static string GetRangeForFormula(int startCol, int endCol, int row)
        {
            return GetColumnLetter(startCol) + row + ":" + GetColumnLetter(endCol) + row;
        }

        private static string GetRangeForFormula(List<int> cols, int row)
        {
            string result = string.Empty;
            foreach (int col in cols)
            {
                result += GetColumnLetter(col) + row;

                if (col != cols.Last())
                {
                    result += ", ";
                }
            }

            return result;
        }

        private static void SetBorder(Excel.Worksheet sheet, int startRow, int startCol, int stopRow, int stopCol)
        {
            Excel.Range formatRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[stopRow, stopCol]];

            List<Excel.XlBordersIndex> borders = new ()
            {
                Excel.XlBordersIndex.xlEdgeLeft,
                Excel.XlBordersIndex.xlEdgeRight,
                Excel.XlBordersIndex.xlEdgeTop,
                Excel.XlBordersIndex.xlEdgeBottom,
                Excel.XlBordersIndex.xlInsideVertical,
                Excel.XlBordersIndex.xlInsideHorizontal,
            };

            foreach (Excel.XlBordersIndex border in borders)
            {
                formatRange.Borders[border].LineStyle = Excel.XlLineStyle.xlContinuous;
                formatRange.Borders[border].Weight = Excel.XlBorderWeight.xlThin;
                formatRange.Borders[border].Color = Excel.XlColorIndex.xlColorIndexAutomatic;
                formatRange.Borders[border].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            }
        }

        private static void SetOutsideBorder(Excel.Worksheet sheet, int startRow, int startCol, int stopRow, int stopCol)
        {
            Excel.Range formatRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[stopRow, stopCol]];

            List<Excel.XlBordersIndex> borders = new ()
            {
                Excel.XlBordersIndex.xlEdgeLeft,
                Excel.XlBordersIndex.xlEdgeRight,
                Excel.XlBordersIndex.xlEdgeTop,
                Excel.XlBordersIndex.xlEdgeBottom,
            };

            foreach (Excel.XlBordersIndex border in borders)
            {
                formatRange.Borders[border].LineStyle = Excel.XlLineStyle.xlContinuous;
                formatRange.Borders[border].Weight = Excel.XlBorderWeight.xlThick;
                formatRange.Borders[border].Color = Excel.XlColorIndex.xlColorIndexAutomatic;
                formatRange.Borders[border].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            }
        }

        private static void SetCellOnSheet(
            Excel.Worksheet sheet,
            int row,
            int col,
            string value,
            bool header = false,
            System.Drawing.KnownColor highlightColor = System.Drawing.KnownColor.Transparent)
        {
            sheet.Cells[row, col] = value;

            if (header)
            {
                SetCellAsHeader(sheet, row, col);
            }

            if (highlightColor != System.Drawing.KnownColor.Transparent)
            {
                SetCellHighlight(sheet, row, col, highlightColor);
            }
        }

        private void SetCellOnAllSheets(
            int row,
            int col,
            string value,
            bool header = false,
            System.Drawing.KnownColor highlightColor = System.Drawing.KnownColor.Transparent)
        {
            foreach (Excel.Worksheet sheet in worksheets)
            {
                SetCellOnSheet(sheet, row, col, value, header: header, highlightColor: highlightColor);
            }
        }

        private void MergeCellsOnAllSheets(int startRow, int startCol, int stopRow, int stopCol)
        {
            foreach (Excel.Worksheet sheet in worksheets)
            {
                MergeCells(sheet, startRow, startCol, stopRow, stopCol);
            }
        }

        private Excel.Worksheet AddSheet(Excel.Workbook workbook, string name, Excel.Worksheet after)
        {
            Excel.Worksheet sheet;
            sheet = workbook.Worksheets.Add(After: after);
            sheet.Name = name;
            worksheets.Add(sheet);

            return sheet;
        }

        private void AutoFixAllSheets()
        {
            foreach (Excel.Worksheet sheet in worksheets)
            {
                sheet.Rows.AutoFit();
                sheet.Columns.AutoFit();
            }
        }
    }
}