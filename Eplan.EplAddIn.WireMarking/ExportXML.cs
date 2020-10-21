using Eplan.EplApi.ApplicationFramework;
using System;
using System.Diagnostics;
using Action = Eplan.EplApi.ApplicationFramework.Action;

namespace WireMarking
{

    class ExportXML
    {

        public static void Execute(string xmlExportFileName)
        {
            string config_scheme = "Маркировка проводов для Partex без обратного адреса XML";

            String strAction = "label";
            ActionManager oAMnr = new ActionManager();
            Action oAction = oAMnr.FindAction(strAction);
            if (oAction != null)
            {
                ActionCallingContext ctx = new ActionCallingContext();

                ctx.AddParameter("CONFIGSCHEME", config_scheme);
                ctx.AddParameter("LANGUAGE", "??_??");
                ctx.AddParameter("DESTINATIONFILE", @"$(TMP)\" + xmlExportFileName);
                // ctx.AddParameter("RECREPEAT", "1");
                // ctx.AddParameter("TASKREPEAT", "1");               

                bool bRet = oAction.Execute(ctx);
                if (bRet == false)
                {
                    DoWireMarking.MassageHandler(strAction);
                }
            }

            Debug.WriteLine(@"-----------------");
            Debug.WriteLine(@"-------NEW-------");
            Debug.WriteLine(@"-----------------");
            Debug.WriteLine(@"$(TMP)\" + xmlExportFileName);

        }
    }
}
