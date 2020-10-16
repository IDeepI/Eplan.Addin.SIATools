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
            OurMenu.AddMainMenu("SIA", Menu.MainMenuName.eMainMenuUtilities, "Маркировать", "WireMarking",
                "Вывод маркировки", 1);
            return true;
        }

        public bool OnExit()
        {
            return true;
        }
    }
}