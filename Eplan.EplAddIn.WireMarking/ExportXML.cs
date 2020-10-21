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
            // Scheme of marking export
            string config_scheme = "Маркировка проводов для Partex без обратного адреса XML";
            // Action
            string strAction = "label";
            ActionManager oAMnr = new ActionManager();
            Action oAction = oAMnr.FindAction(strAction);
            if (oAction != null)
            {
                // Action properties
                ActionCallingContext ctx = new ActionCallingContext();

                ctx.AddParameter("CONFIGSCHEME", config_scheme);
                ctx.AddParameter("LANGUAGE", "??_??");
                ctx.AddParameter("DESTINATIONFILE", @"$(TMP)\" + xmlExportFileName);              

                bool bRet = oAction.Execute(ctx);
                if (bRet == false)
                {
                    DoWireMarking.MassageHandler(strAction);
                }
            }
            // Debug info
            Debug.WriteLine(@"-----------------");
            Debug.WriteLine(@"-------NEW-------");
            Debug.WriteLine(@"-----------------");
            Debug.WriteLine(@"$(TMP)\" + xmlExportFileName);

        }
    }
}
