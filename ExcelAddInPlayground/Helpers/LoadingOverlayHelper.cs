using System;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddInPlayground
{
    public class LoadingOverlayHelper
    {
        private static Excel.Shape _loadingText = null;

        public static void StartLoading()
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Window activeWindow = app.ActiveWindow;
            Excel.Worksheet currentSheet = app.ActiveSheet;
            Excel.Range visibleRange = activeWindow.VisibleRange;

            System.Drawing.Size textSize = System.Windows.Forms.TextRenderer.MeasureText("Loading please wait ...", new System.Drawing.Font("Arial", 22));

            Excel.Shapes shapes = currentSheet.Shapes;

            float left = Convert.ToSingle(visibleRange.Width / 2) - textSize.Width / 2;
            float top = Convert.ToSingle(visibleRange.Height / 2) - (textSize.Height / 2 + 40);

            if (visibleRange.Row != 1)
            {
                // Add any top margin if there is a scroll top
                Excel.Range invisibleRangeOnTop = currentSheet.Range[currentSheet.Cells[1, 1], currentSheet.Cells[visibleRange.Row - 1, 1]];
                top += invisibleRangeOnTop.Height;
            }

            if (visibleRange.Column != 1)
            {
                // Add any left margin if there is a scroll to right
                Excel.Range invisibleRangeAside = currentSheet.Range[currentSheet.Cells[1, 1], currentSheet.Cells[1, visibleRange.Column - 1]];
                left += invisibleRangeAside.Width;
            }

            _loadingText = shapes.AddTextEffect(Office.MsoPresetTextEffect.msoTextEffect16,
                "Loading please wait ...",
                "Arial",
                22,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoFalse,
                left,
                top
            );

            app.Interactive = false;
        }

        public static void StopLoading()
        {
            if (_loadingText != null)
            {
                _loadingText.Delete();
            }

            Globals.ThisAddIn.Application.Interactive = true;
        }
    }
}
