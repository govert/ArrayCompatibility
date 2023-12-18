using System;
using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;
using Microsoft.Office.Interop.Excel;

namespace ArrayCompatibility
{
    // This class defines a few test functions that can be used to explore the automatic array resizing.
    public static class ResizeTestFunctions
    {
        // Just returns an array of the given size
        public static object[,] dnaMakeArray(int rows, int columns)
        {
            object[,] result = new object[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    result[i, j] = i + j;
                }
            }

            return result;
        }

        // Makes an array, but automatically resizes the result
        public static object dnaMakeArrayAndResize(int rows, int columns, string unused, string unusedtoo)
        {
            object[,] result = dnaMakeArray(rows, columns);
            return ArrayResizer.dnaResize(result);

            // Can also call Resize via Excel - so if the Resize add-in is not part of this code, it should still work
            // (though calling direct is better for large arrays - it prevents extra marshaling).
            // return XlCall.Excel(XlCall.xlUDF, "Resize", result);
        }
    }

    public class ArrayResizer
    {
        // This flag controls whether we convert to resized CSE formulas into dynamic arrays
        static bool _convertResizeToDynamic = true;

        public static object dnaResize(object[,] results)
        {
            var caller = Excel(xlfCaller) as ExcelReference;
            if (caller == null)
            {
                return results;
            }

            int rows = results.GetLength(0);
            int columns = results.GetLength(1);

            if (rows == 0 || columns == 0)
            {
                // Empty array - just return
                return results;
            }

            // For dynamic-array aware Excel
            if (UtilityFunctions.dnaSupportsDynamicArrays())
            {
                if (!_convertResizeToDynamic)
                {
                    // We don't want to convert to dynamic arrays
                    // - just return the result which might be a CSE formula or a Dynamic Array already
                    return results;
                }

                // Check if we have a single cell formula
                // If so, we don't only need to convert to dynamic array if it is a CSE array formula
                if (caller.RowFirst == caller.RowLast &&
                    caller.ColumnFirst == caller.ColumnLast)
                {
                    if (!IsFormulaArray(caller))
                    {
                        // Not a CSE array formula - just return the result - dynamic array will sort out
                        return results;
                    }
                }

                // We have a CSE array formula (single or multi-cell)
                // - convert to dynamic array by calling resize with the cancel flag
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    // (Will trigger a recalc by writing formula)
                    DoResizeCOM(caller, cancelResize: true);
                });
                return results;
            }

            // If we get here, we are running in pre-Dynamic-Array Excel
            if ((caller.RowLast - caller.RowFirst + 1 == rows) &&
                (caller.ColumnLast - caller.ColumnFirst + 1 == columns))
            {
                // Size is already OK - just return result
                return results;
            }

            var rowLast = caller.RowFirst + rows - 1;
            var columnLast = caller.ColumnFirst + columns - 1;

            // Check for the sheet limits
            if (rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 ||
                columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1)
            {
                // Can't resize - goes beyond the end of the sheet - just return #VALUE
                // (Can't give message here, or change cells)
                return ExcelError.ExcelErrorValue;
            }

            // TODO: Add some kind of guard for ever-changing result?
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // Create a reference of the right size
                var target = new ExcelReference(caller.RowFirst, rowLast, caller.ColumnFirst, columnLast, caller.SheetId);
                DoResizeCOM(target); // Will trigger a recalc by writing formula
            });

            // Return the whole array even if we plan to resize - to prevent flashing #N/A
            return results;
        }

        // CancelResize means there is an array formula in target, we want to convert it to a non-array formula in the first cell.
        static void DoResizeCOM(ExcelReference target, bool cancelResize = false)
        {
            string wbSheetName = null;
            Range firstCell = null;
            string formula = null;

            try
            {
                var xlApp = ExcelDnaUtil.Application as Application;
                xlApp.DisplayAlerts = false;

                wbSheetName = Excel(xlSheetNm, target) as string;
                int index = wbSheetName.LastIndexOf(']');
                var wbName = wbSheetName.Substring(1, index - 1);
                var sheetName = wbSheetName.Substring(index + 1);
                var ws = xlApp.Workbooks[wbName].Sheets[sheetName] as Worksheet;
                var targetRange = xlApp.Range[
                    ws.Cells[target.RowFirst + 1, target.ColumnFirst + 1],
                    ws.Cells[target.RowLast + 1, target.ColumnLast + 1]] as Range;

                firstCell = targetRange.Cells[1, 1];
                formula = firstCell.Formula;
                if (firstCell.HasArray)
                    firstCell.CurrentArray.ClearContents();
                else
                    firstCell.ClearContents();

                if (cancelResize)
                    ((dynamic)firstCell).Formula2 = formula;
                else
                    targetRange.FormulaArray = formula;
            }
            catch (Exception ex)
            {
                Excel(xlcAlert, $"Cannot resize array formula at {wbSheetName}!{firstCell?.Address} - result might overlap another array.\r\n({ex.Message})");
                firstCell.Value = "'" + formula;
            }
        }

        static bool IsFormulaArray(ExcelReference target)
        {
            // Easy and fast using the C API, but requires the registered function to be IsMacroType=true
            // return (bool)Excel(xlfGetCell, 49, caller);

            // Slow COM approach which we wouldn't really want in a UDF
            var xlApp = ExcelDnaUtil.Application as Application;
            var wbSheetName = Excel(xlSheetNm, target) as string;
            int index = wbSheetName.LastIndexOf(']');
            var wbName = wbSheetName.Substring(1, index - 1);
            var sheetName = wbSheetName.Substring(index + 1);
            var ws = xlApp.Workbooks[wbName].Sheets[sheetName] as Worksheet;
            var firstCell = xlApp.Range[
                ws.Cells[target.RowFirst + 1, target.ColumnFirst + 1],
                ws.Cells[target.RowFirst + 1, target.ColumnFirst + 1]] as Range;

            return firstCell.HasArray;
        }
    }

    public static class UtilityFunctions
    {
        static bool? _supportsDynamicArrays;
        [ExcelFunction(IsHidden = true)]
        public static bool dnaSupportsDynamicArrays()
        {
            if (!_supportsDynamicArrays.HasValue)
            {
                try
                {
                    var result = XlCall.Excel(614, new object[] { 1 }, new object[] { true });
                    _supportsDynamicArrays = true;
                }
                catch
                {
                    _supportsDynamicArrays = false;
                }
            }
            return _supportsDynamicArrays.Value;
        }
    }

}
