using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using System;
using System.Diagnostics;
using Action = Eplan.EplApi.ApplicationFramework.Action;

namespace WireMarking
{

    class ExportXML
    {



        public static void Execute(string xmlExportFileName)
        {
            string config_scheme = "Маркировка проводов для Partex без обратного адреса";         

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
                if (bRet)
                {
                    new Decider().Decide(EnumDecisionType.eOkDecision, "The Action " + strAction + " ended successfully!", "", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);
                }
                else
                {
                    new Decider().Decide(EnumDecisionType.eOkDecision, "The Action " + strAction + " ended with errors!", "", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);
                }
            }

            Debug.WriteLine(@"-----------------");
            Debug.WriteLine(@"-------NEW-------");
            Debug.WriteLine(@"-----------------");
            Debug.WriteLine(@"$(TMP)\" + xmlExportFileName);

            /* m_pClient.SynchronousMode = true;
             CallingContext oCallingContext = new CallingContext();

             oCallingContext.Set("PROJECTNAME", sProjectFileName);
             oCallingContext.Set("CONFIGSCHEME", config_scheme);
             oCallingContext.Set("LANGUAGE", "??_??");
             oCallingContext.Set("DESTINATIONFILE", @"$(TMP)\" + xmlExportFileName);
             oCallingContext.Set("RECREPEAT", "1");
             oCallingContext.Set("TASKREPEAT", "1");

             m_pClient.ExecuteAction($"label", ref oCallingContext);

             
             Debug.WriteLine(oCallingContext.Message);*/

            //Console.ReadKey();


        }
    }
}
