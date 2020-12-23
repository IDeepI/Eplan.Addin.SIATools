using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;

namespace SIATools.ExportToPdfAndDwg
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
            progress.SetTitle("DWG/PDF export");
            progress.ShowImmediately();
            progress.BeginPart(25.0, "ChangeFontType to GOST Type AU : ");
            try
            {
                ChangeDrawMode(2);
                ChangeFontType(CurrentProject, "GOST Type AU");
                progress.EndPart();
                progress.BeginPart(25.0, "Export to DWG : ");
                ExportToDwg();
            }
            catch (Exception ex)
            {
                progress.EndPart(true);
                DoWireMarking.DoWireMarking.ErrorHandler("Export to DWG", ex);
                return false;
            }


            progress.EndPart();
            progress.BeginPart(25.0, "ChangeFontType to GOST type A : ");
            progress.Step(1);
            try
            {
                ChangeFontType(CurrentProject, "GOST type A");
                progress.EndPart();
                progress.BeginPart(25.0, "Export to PDF : ");
                ExportToPdf();
                ChangeDrawMode(3);
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
        /// How to draw point of connection
        /// </summary>
        /// <param name="drawMode"> 3 - default. 2 - for printing</param>
        private void ChangeDrawMode(int drawMode)
        {
            Eplan.EplApi.DataModel.ProjectManager oProjectManager = new Eplan.EplApi.DataModel.ProjectManager();
            Project oProject = oProjectManager.CurrentProject;
            Eplan.EplApi.DataModel.ProjectSettings projectSettings = new Eplan.EplApi.DataModel.ProjectSettings(oProject);

            string befor = projectSettings.GetExpandedStringSetting("TrDMProject.Wiring", 0);

            projectSettings.SetNumericSetting("TrDMProject.Wiring", drawMode, 0);

            //DoWireMarking.DoWireMarking.MassageHandler($"Befor { befor }\nAfter { drawMode }");
        }

        /// <summary>
        /// Export to pdf with filter "Для печати" and scheme "SIA"
        /// </summary>
        private void ExportToPdf()
        {
            SelectionSet Set = new SelectionSet();
            Project CurrentProject = Set.GetCurrentProject(true);
            string ProjectName = CurrentProject.ProjectName;
            // Scheme of marking export
            string exportType = "PDFPROJECTSCHEME";
            string exportScheme = "SIA";
            string exportFileName = $"d:\\Work\\PDF\\{ ProjectName }_{ DateTime.Now.Year }.{ DateTime.Now.Month }.{ DateTime.Now.Day }.pdf";
            // Action
            string strAction = "export";

            // Export a project in pdf format            

            ActionManager oAMnr = new ActionManager();
            Eplan.EplApi.ApplicationFramework.Action oAction = oAMnr.FindAction(strAction);
            if (oAction != null)
            {
                // Action properties
                ActionCallingContext ctx = new ActionCallingContext();

                ctx.AddParameter("TYPE", exportType);
                ctx.AddParameter("EXPORTSCHEME", exportScheme);
                ctx.AddParameter("EXPORTFILE", exportFileName);
                // ctx.AddParameter("USEPAGEFILTER", "1");

                bool bRet = oAction.Execute(ctx);
                if (bRet == false)
                {
                    DoWireMarking.DoWireMarking.MassageHandler("Error in Action - ExportToPdf");
                }
            }

        }
        /// <summary>
        /// Export to dwg with filter "Для печати" and scheme "SIA DWG"
        /// </summary>
        private void ExportToDwg()
        {
            // Scheme of marking export
            string exportType = "DWGPROJECT";
            //string exportScheme = "SIA";
            string exportPath = @"d:\Work\DWG\";
            // Action
            string strAction = "export";

            // Export a project in DXF / DWG format
            ActionManager oAMnr = new ActionManager();
            Eplan.EplApi.ApplicationFramework.Action oAction = oAMnr.FindAction(strAction);
            if (oAction != null)
            {
                // Action properties
                ActionCallingContext ctx = new ActionCallingContext();

                ctx.AddParameter("TYPE", exportType);
                // ctx.AddParameter("EXPORTSCHEME", exportScheme);            
                ctx.AddParameter("DESTINATIONPATH", exportPath);
                //ctx.AddParameter("USEPAGEFILTER", "1");

                bool bRet = oAction.Execute(ctx);
                if (bRet == false)
                {
                    DoWireMarking.DoWireMarking.MassageHandler("Error in Action - ExportToDwg");

                    DoWireMarking.DoWireMarking.MassageHandler(ctx.ToString());
                    DoWireMarking.DoWireMarking.MassageHandler(ctx.GetParameters().ToString());
                    DoWireMarking.DoWireMarking.MassageHandler(ctx.GetStrings().ToString());

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
