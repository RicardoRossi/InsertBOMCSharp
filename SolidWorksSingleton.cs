using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace InsertBOMCSharp
{
    internal class SolidWorksSingleton
    {
        private static SldWorks swApp;

        private SolidWorksSingleton()
        {

        }

        internal static SldWorks GetApplication()
        {

            if (swApp == null)
            {
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    return swApp;
                }
                catch (Exception)
                {
                    swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
                    swApp.Visible = true;
                    return swApp;
                }

            }
            return swApp;
        }

        internal static void Dispose()
        {
            if (swApp != null)
            {
                swApp.ExitApp();
                swApp = null;

            }
        }

    }
}
