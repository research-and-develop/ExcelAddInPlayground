using ExcelAddInPlayground.Forms;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddInPlayground
{
    public class ContextMenuHandler
    {
        private List<Excel.Shape> _textEffects;
        private Office.CommandBar _contextMenu;
        private Excel.Application _app;

        public ContextMenuHandler()
        {
            _app = Globals.ThisAddIn.Application;
            _contextMenu = _app.CommandBars["Cell"];
            _textEffects = new List<Excel.Shape>();
        }

        public void ResetMenu()
        {
            _contextMenu.Reset();
        }

        public void InitializelMenu()
        {
            // Reset before initializing
            ResetMenu();

            //=================================
            // Sample Popup Menu
            Office.CommandBarPopup popupMenu = (Office.CommandBarPopup)_contextMenu.Controls
                .Add(Office.MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            popupMenu.Caption = "Sample Popup Menu";

            Office.CommandBarButton printTextEffectsMenuItem = (Office.CommandBarButton)popupMenu.Controls
                .Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            printTextEffectsMenuItem.Caption = "Print All Text Effects";
            printTextEffectsMenuItem.FaceId = 0609;
            printTextEffectsMenuItem.Click += PrintTextEffectsMenuItem_Click;

            Office.CommandBarButton clearTextEffectsMenuItem = (Office.CommandBarButton)popupMenu.Controls
                .Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            clearTextEffectsMenuItem.Caption = "Clear Printed Text Effects";
            clearTextEffectsMenuItem.FaceId = 0330;
            clearTextEffectsMenuItem.Click += ClearTextEffectsMenuItem_Click;



            //=================================
            // Sample Long Processing popup menu 
            Office.CommandBarPopup longProcessingMenu = (Office.CommandBarPopup)_contextMenu.Controls
                .Add(Office.MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            longProcessingMenu.Caption = "Sample Long Processing";

            Office.CommandBarButton showLoadingFormMenuItem = (Office.CommandBarButton)longProcessingMenu.Controls
                .Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            showLoadingFormMenuItem.Caption = "With Loading Form";
            showLoadingFormMenuItem.FaceId = 0302;
            showLoadingFormMenuItem.Click += ShowLoadingFormMenuItem_Click;

            Office.CommandBarButton showLoadingTextMenuItem = (Office.CommandBarButton)longProcessingMenu.Controls
                .Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            showLoadingTextMenuItem.Caption = "With Loading Text";
            showLoadingTextMenuItem.FaceId = 0253;
            showLoadingTextMenuItem.Click += ShowLoadingTextMenuItem_Click; ;
        }

        private void ShowLoadingTextMenuItem_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            }

            var context = TaskScheduler.FromCurrentSynchronizationContext();
            var task = new Task(() =>
            {
                Globals.ThisAddIn.Application.Interactive = false;
                LoadingOverlayHelper.StartLoading();
            });

            Task continuation = task.ContinueWith(t =>
            {

                Globals.ThisAddIn.Application.StatusBar = "Doing some long processing task ...";

                Thread.Sleep(4000);

                LoadingOverlayHelper.StopLoading();

                Globals.ThisAddIn.Application.StatusBar = "Ready";

                Globals.ThisAddIn.Application.Interactive = true;

            }, CancellationToken.None, TaskContinuationOptions.OnlyOnRanToCompletion, context);

            task.Start();
        }

        private void ShowLoadingFormMenuItem_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FormLoading loader = new FormLoading();
            loader.StartPosition = FormStartPosition.CenterScreen;
            loader.Show();
            
            // This is a pretty important part because on the FromCurrentSynchronizationContext() call you can get Exception
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            }

            var context = TaskScheduler.FromCurrentSynchronizationContext();
            var task = new Task(() =>
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                app.Interactive = false;

                for (int i = 0; i < 500; i++)
                {
                    app.StatusBar = string.Format("Iteration {0}", i);
                    Thread.Sleep(5);
                }

                // This code is used to call Close() method Thread safe otherwise you get Cross-thread operation not valid Exception
                ThreadSafeHelper.InvokeControlMethodThreadSafe(loader, () =>
                {
                    loader.Close();
                });

                app.StatusBar = "Ready";
                app.Interactive = true;
            });

            task.Start();
        }

        private void ClearTextEffectsMenuItem_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            foreach (Excel.Shape textEffect in _textEffects)
            {
                textEffect.Delete();
            }

            _textEffects.Clear();
        }

        private void PrintTextEffectsMenuItem_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            System.Drawing.Size size = TextRenderer.MeasureText("Text Effect  ", new System.Drawing.Font("Arial", 20));

            Excel.Worksheet sheet = _app.ActiveWorkbook.ActiveSheet;
            Excel.Shapes shapes = sheet.Shapes;

            int top = 0;
            for (int i = 0; i < 50; i++)
            {
                Excel.Shape shape = shapes.AddTextEffect((Office.MsoPresetTextEffect)i,
                    string.Format("Text Effect {0}", i + 1),
                    "Arial",
                    20,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoFalse,
                    Convert.ToSingle(_app.ActiveWindow.VisibleRange.Width / 2) - size.Width / 2,
                    top
                );

                top += 30;

                _textEffects.Add(shape);
            }

            return;
        }

    }
}
