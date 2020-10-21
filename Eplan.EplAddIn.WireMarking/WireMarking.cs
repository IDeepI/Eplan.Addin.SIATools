using Eplan.Addin.WireMarking;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace WireMarking
{
    public class DoWireMarking : IEplAction
    {
        // Temp file
        public static string xmlExportFileName = "TMP_XMLWireData.xml";
        /// List of XML objects
        public static List<EplanLabellingDocumentPageLine> listOfLines;

        /// Registr Action under the name ""
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "WireMarking";
            Ordinal = 20;
            return true;
        }
        /// <summary>
        /// Execute Action
        /// </summary>
        /// <param name="oActionCallingContext"></param>
        /// <returns></returns>
        public bool Execute(ActionCallingContext oActionCallingContext)
        {
            SelectionSet Set = new SelectionSet();
            Project CurrentProject = Set.GetCurrentProject(true);
            string ProjectName = CurrentProject.ProjectName;
            Debug.WriteLine(ProjectName);
            string xlsFileName = Path.GetDirectoryName(CurrentProject.ProjectFullName);
            xlsFileName = Path.Combine(xlsFileName, "Marking.xls");
            try
            {
                // Executing Action "label"
                ExportXML.Execute(xmlExportFileName);
            }
            catch (Exception ex)
            {
                ErrorHandler("ExportXML", ex);
                return false;
            }
            try
            {
                // Getting object from XML
                ParseXMLWireFile();
            }
            catch (Exception ex)
            {
                ErrorHandler("ParseXMLWireFile", ex);
                return false;
            }
            try
            {
                // Export to excel
                // Creating *.xls file
                ExportToExcel.Execute(listOfLines, xlsFileName);
            }
            catch (Exception ex)
            {
                ErrorHandler("ExportToExcel", ex);
                return false;
            }
            return true;
        }
        /// <summary>
        /// Show message in Eplan
        /// </summary>
        /// <param name="errorMessage"></param>
        internal static void MassageHandler(string errorMessage)
        {
            new Decider().Decide(EnumDecisionType.eOkDecision, errorMessage, "", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);
        }
        /// <summary>
        /// Show error in Eplan
        /// </summary>
        /// <param name="actionName"></param>
        /// <param name="exception"></param>
        internal static void ErrorHandler(string actionName, Exception exception)
        {
            new Decider().Decide(EnumDecisionType.eOkDecision, "The Action " + actionName + " ended with errors! " + exception.Message, "", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);
        }


        public void GetActionProperties(ref ActionProperties actionProperties)
        {
        }
        /// <summary>
        /// Extract data from xml and serialize it to objects
        /// </summary>
        private static void ParseXMLWireFile()
        {
            // объект для сериализации
            EplanLabelling eplanLabelling = new EplanLabelling();
            Debug.WriteLine($"Объект создан : {System.IO.Path.GetTempPath() + xmlExportFileName}");

            // передаем в конструктор тип класса
            XmlSerializer formatter = new XmlSerializer(typeof(EplanLabelling));

            // десериализация
            using (FileStream fs = new FileStream(Path.GetTempPath() + xmlExportFileName, FileMode.OpenOrCreate))
            {
                EplanLabelling newEplanLabelling = (EplanLabelling)formatter.Deserialize(fs);
                Debug.WriteLine("Объект десериализован");

                listOfLines = newEplanLabelling.Document.Page.Line.ToList();

                // Call Sort on the list. This will use the
                // default comparer, which is the Compare method
                // implemented on Part.
                listOfLines.Sort();

                foreach (var line in listOfLines)
                {
                    foreach (var property in line.Label.Property)
                    {
                        Debug.Write($"{property.PropertyValue}\t : \t");
                    }
                    Debug.WriteLine("");
                }


            }

        }


    }
}