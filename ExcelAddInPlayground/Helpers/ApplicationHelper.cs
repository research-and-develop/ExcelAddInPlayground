using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddInPlayground
{
    public class ApplicationHelper
    {
        public static bool WorkbookHasModule(Excel.Workbook wb, string moduleName)
        {
            bool result = false;

            VBProject project = wb.VBProject;

            foreach (VBComponent component in project.VBComponents)
            {
                if (component.CodeModule.Name == moduleName)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        public static void WorkbookGenerateModule(Excel.Workbook wb, string moduleName, string moduleContent)
        {
            VBProject prj = wb.VBProject;
            VBComponent xlModule = prj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);

            xlModule.Name = moduleName;

            xlModule.CodeModule.AddFromString(moduleContent);
        }

        public static Excel.Worksheet GetSheet(string caption)
        {
            Excel.Worksheet result = null;

            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                if (sheet.Name == caption)
                {
                    result = sheet;
                    break;
                }
            }

            return result;
        }

        public static bool RangeNameExists(string rangeName)
        {
            bool result = false;
            Excel.Application app = Globals.ThisAddIn.Application;

            // first check in application Names
            foreach (Excel.Name appName in app.Application.Names)
            {
                if (appName.NameLocal == rangeName)
                {
                    result = true;
                    break;
                }
            }

            // If didnt found in application names try searching in sheet names
            if (!result)
            {
                foreach (Excel.Worksheet sheet in app.ActiveWorkbook.Worksheets)
                {
                    foreach (Excel.Name sheetName in sheet.Names)
                    {
                        if (sheetName.NameLocal == rangeName)
                        {
                            result = true;
                            break;
                        }
                    }
                    if (result)
                    {
                        break;
                    }
                }
            }

            return result;
        }

        public static Excel.Range RangeGet(string rangeName)
        {
            Excel.Range result = null;
            Excel.Application app = Globals.ThisAddIn.Application;

            // first check in application Names
            foreach (Excel.Name appName in app.Application.Names)
            {
                if (appName.NameLocal == rangeName)
                {
                    result = appName.RefersToRange;
                    break;
                }
            }

            // If didn't found in application names try searching in sheet names
            if (result == null)
            {
                foreach (Excel.Worksheet sheet in app.ActiveWorkbook.Worksheets)
                {
                    foreach (Excel.Name sheetName in sheet.Names)
                    {
                        if (sheetName.NameLocal == rangeName)
                        {
                            result = sheetName.RefersToRange;
                            break;
                        }
                    }
                    if (result != null)
                    {
                        break;
                    }
                }
            }

            return result;
        }

        public static object RangeGetValue(string rangeName)
        {
            Excel.Range range = RangeGet(rangeName);
            if (range == null)
            {
                return null;
            }
            else
            {
                return range.Value;
            }
        }
    }
}
