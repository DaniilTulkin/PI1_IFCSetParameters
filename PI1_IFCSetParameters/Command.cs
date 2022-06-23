using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PI1_CORE.Helpers;
using System.Collections.Generic;

namespace PI1_IFCSetParameters
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    // Start command class.
    public class Command : IExternalCommand
    {
        /// <summary>
        /// Overload this method to implement and external command within Revit.
        /// </summary>
        /// <param name="commandData">An ExternalCommandData object which contains reference to Application and View
        /// needed by external command.</param>
        /// <param name="message">Error message can be returned by external command. This will be displayed only if the command status
        /// was "Failed".  There is a limit of 1023 characters for this message; strings longer than this will be truncated.</param>
        /// <param name="elements">Element set indicating problem elements to display in the failure dialog.  This will be used
        /// only if the command status was "Failed".</param>
        /// <returns>
        /// The result indicates if the execution fails, succeeds, or was canceled by user. If it does not
        /// succeed, Revit will undo any changes made by the external command.
        /// </returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var IFCRebars = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Rebar);

            using (Transaction t = new Transaction(doc, "Запись параметров в IFC арматуру"))
            {
                t.Start();

                foreach (FamilyInstance IFCRebar in IFCRebars)
                {
                    Element host = null;
                    try
                    {
                        host = IFCRebar.Host;
                    }
                    catch { }

                    if (host != null)
                    {
                        Category hostCategory = host.Category;
                        string hostCategoryName = HostName.GetHostName(hostCategory);
                        string hostMark = host.get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString();

                        if (hostCategoryName != null)
                        {
                            IFCRebar.LookupParameter("ADSK_Категория основы").Set(hostCategoryName);
                            IFCRebar.LookupParameter("ADSK_Метка основы").Set(hostMark);
                        }
                        else
                        {
                            IFCRebar.LookupParameter("ADSK_Категория основы").Set(@"<неопредленная категория>");
                            IFCRebar.LookupParameter("ADSK_Метка основы").Set(@"<неопредленная категория>");
                        }
                    }
                    else
                    {
                        IFCRebar.LookupParameter("ADSK_Категория основы").Set(@"<не связано>");
                        IFCRebar.LookupParameter("ADSK_Метка основы").Set(@"<не связано>");
                    }
                }

                t.Commit();
            }
                
            return Result.Succeeded;
        }

        /// <summary>
        /// Gets the path of the current command.
        /// </summary>
        /// <returns></returns>
        public static string GetPath()
        {
            return typeof(Command).Namespace + "." + nameof(Command);
        }
    }
}
