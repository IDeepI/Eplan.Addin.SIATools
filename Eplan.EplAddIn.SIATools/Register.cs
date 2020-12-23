using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;

namespace WireMarking
{
    public class AddInModule : IEplAddIn
    {
        public bool OnRegister(ref bool bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }

        public bool OnUnregister()
        {
            return true;
        }

        public bool OnInit()
        {
            return true;
        }

        public bool OnInitGui()
        {
            Menu OurMenu = new Menu();
            //publicAddMainMenu(string strMainMenuText, string strMainMenuTextReference, string strMenuName, string strActionName, string strStatusText, int nInsertPosition);
            uint nCommandId = OurMenu.AddMainMenu("SIA", Menu.MainMenuName.eMainMenuUtilities, "Маркировать", "WireMarking", "Вывод маркировки", 1);
            //OurMenu.AddMenuItem("Test", "ExportToPdfAndDwg");

            //public uint AddMenuItem(string strText, string strActionName, string strStatusText, uint nCommandId, int nInsertPosition, bool bSeparatorBefore, bool bSeparatorBehind);
            ///OurMenu.AddMenuItem("Печать", "ExportToPdfAndDwg", "strStatusText", 1, 1, false, false);
           
            //public uint AddMenuItem(string strText, string strActionName, string strStatusText, uint nCommandId, int nInsertPosition, bool bSeparatorBefore, bool bSeparatorBehind);
             OurMenu.AddMenuItem("Печать", "ExportToPdfAndDwg", "Экспорт в PDF/DWG", nCommandId, 1, false, false);

           // OurMenu.AddMainMenu("Печать", Menu.MainMenuName.eMainMenuUtilities, "Печать", "ExportToPdfAndDwg", "Вывод в PDF и DWG", 1);

            return true;
        }

        public bool OnExit()
        {
            return true;
        }
    }
}