using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInPlayground
{

    [ComVisible(true)]
    public interface ICSharpFunctionsProxy
    {
        void MessageBoxFromCSharp(string message);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class CSharpFunctionsProxy : ICSharpFunctionsProxy
    {
        private Excel.Application _app;

        public CSharpFunctionsProxy()
        {
            _app = Globals.ThisAddIn.Application;
        }

        public void MessageBoxFromCSharp(string message)
        {
            System.Windows.Forms.MessageBox.Show(message);
        }


    }

}
