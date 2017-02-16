using System;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelAddInPlayground
{
    public class ThreadSafeHelper
    {
        private delegate void SetControlPropertyThreadSafeDelegate(Control control, string propertyName, object propertyValue);
        public static void SetControlPropertyThreadSafe(Control control, string propertyName, object propertyValue)
        {
            if (control.InvokeRequired)
            {
                control.Invoke(new SetControlPropertyThreadSafeDelegate(SetControlPropertyThreadSafe), new object[] { control, propertyName, propertyValue });
            }
            else
            {
                control.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, control, new object[] { propertyValue });
            }
        }

        public static void InvokeControlMethodThreadSafe(Control control, Action action)
        {
            control.Invoke(new MethodInvoker(action), null);
        }
    }
}
