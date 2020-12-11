using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;

namespace WireMarking.ExportToPdfAndDwg
{
    class ExportToPdfAndDwg : IEplAction
    {
        public bool Execute(ActionCallingContext oActionCallingContext)
        {
            SelectionSet Set = new SelectionSet();
            Project CurrentProject = Set.GetCurrentProject(true);
            string ProjectName = CurrentProject.ProjectName;
            //DoWireMarking.DoWireMarking.MassageHandler(ProjectName);

            // Show ProgressBar
            Progress progress = new Progress("SimpleProgress");
            progress.SetAllowCancel(true);
            progress.SetAskOnCancel(true);
            progress.SetTitle("Wire mark export");
            progress.ShowImmediately();
            progress.BeginPart(25.0, "ChangeFontType to GOST Type AU : ");
            try
            {
                ChangeFontType(CurrentProject, "GOST Type AU");
                progress.EndPart();
                progress.BeginPart(25.0, "Export to DWG : ");
                ExportToDwg();               
            }
            catch (Exception ex)
            {
                DoWireMarking.DoWireMarking.ErrorHandler("Export to DWG", ex);
                return false;
            }

            progress.EndPart();
            progress.BeginPart(25.0, "ChangeFontType to GOST type A : ");

            try
            {
                ChangeFontType(CurrentProject, "GOST type A");
                progress.EndPart();
                progress.BeginPart(25.0, "Export to PDF : ");
                ExportToPdf();
            }
            catch (Exception ex)
            {
                DoWireMarking.DoWireMarking.ErrorHandler("Export to PDF", ex);
                return false;
            }
            finally
            {
                progress.EndPart(true);
            }

            return true;
        }
        /// <summary>
        /// Export to pdf with filter "Для печати" and scheme "SIA"
        /// </summary>
        private void ExportToPdf()
        {
            // Scheme of marking export
            string exportType = "PDFPROJECTSCHEME";
            string exportScheme = "SIA";
            string exportFileName = "ESS_Sample_Project";
            // Action
            string strAction = "export";

            // Export a project in pdf format

            // export / TYPE:PDFPROJECTSCHEME / PROJECTNAME:C:\Projects\EPLAN\ESS_Sample_Project.elk / EXPORTFILE:C:\ESS_Sample_Project.pdf / EXPORTSCHEME:myScheme

            ActionManager oAMnr = new ActionManager();
            Eplan.EplApi.ApplicationFramework.Action oAction = oAMnr.FindAction(strAction);
            if (oAction != null)
            {
                // Action properties
                ActionCallingContext ctx = new ActionCallingContext();

                ctx.AddParameter("TYPE", exportType);
                ctx.AddParameter("EXPORTSCHEME", exportScheme);
                ctx.AddParameter("EXPORTFILE", @"$(TMP)\" + exportFileName);

                bool bRet = oAction.Execute(ctx);
                if (bRet == false)
                {
                    DoWireMarking.DoWireMarking.MassageHandler(strAction);
                }
            }
            
        }
        /// <summary>
        /// Export to dwg with filter "Для печати" and scheme "SIA DWG"
        /// </summary>
        private void ExportToDwg()
        {
            // Scheme of marking export
            string exportType = "DXFPROJECT";
            //string exportScheme = "SIA";
            string exportPath = "";
            // Action
            string strAction = "export";

            // Export a project in DXF / DWG format

            // export / TYPE:DXFPROJECT / PROJECTNAME:C:\Projects\EPLAN\ESS_Sample_Project.elk / DESTINATIONPATH:C:\temp

               ActionManager oAMnr = new ActionManager();
            Eplan.EplApi.ApplicationFramework.Action oAction = oAMnr.FindAction(strAction);
            if (oAction != null)
            {
                // Action properties
                ActionCallingContext ctx = new ActionCallingContext();

                ctx.AddParameter("TYPE", exportType);
               // ctx.AddParameter("EXPORTSCHEME", exportScheme);            
                ctx.AddParameter("DESTINATIONPATH", @"$(TMP)\" + exportPath);

                bool bRet = oAction.Execute(ctx);
                if (bRet == false)
                {
                    DoWireMarking.DoWireMarking.MassageHandler(strAction);
                }
            }
        }

        /// <summary>
        /// Change first and second firm font type to selected
        /// </summary>
        /// <param name="font">Selected font type</param>

        private void ChangeFontType(Project currentProject, string font)
        {
            Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
          
            oSettings.SetStringSetting("COMPANY.GedViewer.Fonts", $"??_??@{font};", 0);
            oSettings.SetStringSetting("COMPANY.GedViewer.Fonts", $"??_??@{font};", 1);

            string strTest0 = oSettings.GetStringSetting("COMPANY.GedViewer.Fonts", 0);

            DoWireMarking.DoWireMarking.MassageHandler(strTest0);
        }

        public void GetActionProperties(ref ActionProperties actionProperties)
        {
            throw new NotImplementedException();
        }

        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "ExportToPdfAndDwg";
            //Ordinal = 20;
            return true;
        }
    }
}
