using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;

namespace ScreenPopper
{
    class Program
    {
        static void Main(string[] args)
        {
            //string szPhone = "80055512120000AARONFIT0001";
            string szSanScript = "";
            string szErrors = "";
            int iResult = 0;
            Properties.Settings setting = new ScreenPopper.Properties.Settings();
            if (setting.Properties.Count < 5)
            {
                LogError ("Internal Error - missing properties ...");
                return;
            }
            string Formname = setting.FormName;
            string Windowname = setting.WindownName;
            string script = setting.ScriptToRun;
            string process = setting.GPAppProc;
            string field = setting.FieldToSet;
            int iStart = setting.iStart;
            int iTake = setting.iTake;

            if (args.Length < 1)
            {
                LogError("Please pass in the proper value");
                return;
            }

            //TODO check to see if GP is running. If not run GP
                Process[] dynamics = Process.GetProcessesByName(process);
                
            if (dynamics.Length == 0)
                {
                    LogError("GP Dynamics 10 must be running ...");
                    return;
                }
            string szPhone = args[0];

            // Instantiate a new Dynamics.Application reference to a running GP session
            Dynamics.Application app = new Dynamics.Application();

            // Set product to Cogsdale context
            app.CurrentProductID = 229;

            szSanScript = "open form '" + Formname + "';";

            string LocationID = "";

            if (iTake == 0)
            {

                LocationID = szPhone.Substring(iStart, (szPhone.Length - iStart));
            }
            else
            {
                LocationID = szPhone.Substring(iStart, iTake);
            }

            //szSanScript = "open form UM_Location_Account_4; ";
            //szSanScript += "umLocationID of window Maint_Window of form UM_Location_Account_4 = \"\"" + LocationID + "\"\";";
            //szSanScript += "run script umLocationID of window Maint_Window of form UM_Location_Account_4; ";

            // Call SanScript
            iResult = app.ExecuteSanscript(szSanScript, out szErrors);
            
            if (iResult > 0)
            {
                LogError(szErrors);
                return;
            }

            iResult = app.SetDataValue("'" + field + "' of window '" + Windowname + "' of form '" + Formname + "'", LocationID);
            if (iResult > 0)
            {
                LogError(string.Format("The SetDataValue function failed to run. The values passed in were: Field={0} Window Name={1} Form Name={2} LocationID={3}", 
                    field, Windowname, Formname, LocationID));
                return;
            }

            iResult = app.ExecuteSanscript("run script '" + script + "' of window '" + Windowname + "' of form '" + Formname + "';", out szErrors);
            if (iResult > 0)
            {
                LogError(szErrors);
                return;
            }

            //Console.WriteLine("Press any key to continue ..");
            //Console.ReadKey();

        }

        private static void LogError(string szErrors)
        {
            using (TextWriter tw = new StreamWriter("scrnpop.log", false))
            {
                tw.WriteLine(szErrors);
                tw.Close();
            }
        }
    }
}
