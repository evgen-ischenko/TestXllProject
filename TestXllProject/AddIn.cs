using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace TestXllProject
{
    public class AddIn : IExcelAddIn
    {

        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        [ExcelFunction(description: "Just for test")]
        public static object TestFunction([ExcelArgument (description: "Some string")] string name)
        {
            return "Hello " + name;
        }

    }
}
