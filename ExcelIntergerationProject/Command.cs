using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace ExcelIntergerationProject
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .ToList();

            Dictionary<string, string> roomDic = new Dictionary<string, string>();

            foreach (var room in rooms)
                roomDic.Add(room.Id.ToString(), room.Area.ToString());

            ExcelService excelService = new ExcelService();

            excelService.Export("RoomData",
                                "Id",
                                "Area",
                                roomDic);

            return Result.Succeeded;
        }
    }
}
