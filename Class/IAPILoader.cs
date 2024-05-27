using System;

namespace AppLinkReplace.Class
{
    internal interface IAPILoader
    {
        void Initialize();

        bool InitializeTFlexCADAPI(String loadBomdll);

        void Terminate();

        String TflexVersionApi { get; }
        bool IsInit { get; }
    }
}