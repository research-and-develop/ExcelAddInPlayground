using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInPlayground
{
    public partial class ThisAddIn
    {
        private ContextMenuHandler _contextMenuHandler;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _contextMenuHandler = new ContextMenuHandler();

            Application.SheetBeforeRightClick += Application_SheetBeforeRightClick;
            Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            // Add module with functions if not existing
            if (!ApplicationHelper.WorkbookHasModule(Wb, "CSharpProxy"))
            {
                string moduleContent = string.Empty;

                moduleContent +=
                    "Sub MessageBoxFromCSharp(message As String)" + Environment.NewLine +
                    "    Dim addIn As COMAddIn" + Environment.NewLine +
                    "    Dim proxy As Object" + Environment.NewLine +
                    "    Set addIn = Application.COMAddIns(\"ExcelAddInPlayground\")" + Environment.NewLine +
                    "    Set proxy = addIn.Object" + Environment.NewLine +
                    "    proxy.MessageBoxFromCSharp(message)" + Environment.NewLine +
                    "End Sub" + Environment.NewLine + Environment.NewLine;

                ApplicationHelper.WorkbookGenerateModule(Wb, "CSharpProxy", moduleContent);
            }
        }

        private void Application_SheetBeforeRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            _contextMenuHandler.InitializelMenu();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private CSharpFunctionsProxy _proxy;
        protected override object RequestComAddInAutomationService()
        {
            if (_proxy == null)
            {
                _proxy = new CSharpFunctionsProxy();
            }

            return _proxy;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
