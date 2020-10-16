using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;

namespace WireMarking
{
    public class DoWireMarking : IEplAction
    {
        public static string xmlExportFileName = "TMP_XMLWireData.xml";

        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "WireMarking";
            Ordinal = 20;
            return true;
        }

        public bool Execute(ActionCallingContext oActionCallingContext)
        {
            /*  SelectionSet Set = new SelectionSet();
              Project CurrentProject = Set.GetCurrentProject(true);
              string ProjectName = CurrentProject.ProjectName;
              string ProjectCompanyName = CurrentProject.Properties.PROJ_COMPANYNAME;
              DateTime ProjectCreationDate = CurrentProject.Properties.PROJ_CREATIONDATE;
              MessageBox.Show("Название проекта: " + ProjectName + "\n" + "Название фирмы: " + ProjectCompanyName +
                              "\n" + "Дата создания проекта: " + ProjectCreationDate.ToShortDateString());
             */

            ExportXML.Execute(xmlExportFileName);
            ParseXMLWireFile();


            return true;
        }

       

        public void GetActionProperties(ref ActionProperties actionProperties)
        {
        }

        private static void ParseXMLWireFile()
        {
            // объект для сериализации
            EplanLabelling eplanLabelling = new EplanLabelling();
            Debug.WriteLine($"Объект создан : {System.IO.Path.GetTempPath() + xmlExportFileName}");

            // передаем в конструктор тип класса
            XmlSerializer formatter = new XmlSerializer(typeof(EplanLabelling));

           // Console.ReadLine();
            // десериализация
            using (FileStream fs = new FileStream(Path.GetTempPath() + xmlExportFileName, FileMode.OpenOrCreate))
            {
                EplanLabelling newEplanLabelling = (EplanLabelling)formatter.Deserialize(fs);
                Debug.WriteLine("Объект десериализован");

                var listOfLines = newEplanLabelling.Document.Page.Line.ToList();

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

                ///TODO: Export to excel tamplate
            }
            Console.ReadLine();
        }

    }
}