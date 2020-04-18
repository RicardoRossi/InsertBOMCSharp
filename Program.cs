using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
namespace InsertBOMCSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            SldWorks swApp;
            ModelDoc2 swModel;
            swApp = SolidWorksSingleton.GetApplication();
            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (!(GetBomTable(swModel)))
            {
                InsertBOM(swModel);
            }

        }

        private static bool GetBomTable(ModelDoc2 swModel)
        {
            Feature swFeature;
            //BomFeature swBomFeat;
            swFeature = swModel.FirstFeature();
            while (!(swFeature == null))
            {
                string featureName = swFeature.GetTypeName2();
                if (featureName == "BomFeat")
                {
                    return true;
                }
                swFeature = swFeature.GetNextFeature();
            }
            return false;
        }

        private static void InsertBOM(ModelDoc2 swmodel)
        {
            DrawingDoc swDraw = (DrawingDoc)swmodel;
            View swSheetView;
            View swActiveView;
            swSheetView = swDraw.GetFirstView(); // De fato é a folha, por isso a primeira vista é obtida por GetNextView
            swActiveView = swSheetView.GetNextView();
            swmodel = swActiveView.ReferencedDocument;
            string config = swActiveView.ReferencedConfiguration;

            var bomPart = @"C:\Users\Ricar\Documents\Add-in Ricardo\InsertBOMCSharp\PART.sldbomtbt";
            var bomAssembly = @"C:\Users\Ricar\Documents\Add-in Ricardo\InsertBOMCSharp\ASSEMBLY.sldbomtbt";

            if ((swDocumentTypes_e)swmodel.GetType() == swDocumentTypes_e.swDocASSEMBLY)
            {
                swActiveView.InsertBomTable4(false, 0.413, 0.049, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight,
                   (int)swBomType_e.swBomType_TopLevelOnly, config, bomAssembly, false, (int)swNumberingType_e.swNumberingType_None, false);
            }
            else if ((swDocumentTypes_e)swmodel.GetType() == swDocumentTypes_e.swDocPART)
            {
                swActiveView.InsertBomTable4(false, 0.413, 0.049, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight,
                   (int)swBomType_e.swBomType_PartsOnly, config, bomPart, false, (int)swNumberingType_e.swNumberingType_None, false);
            }
            //Console.ReadKey();
        }


    }
}
