using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DBLibrary;

namespace UpdateFamilyTypes_DataBase
{
    class ProgramUpdateTypes
    {
        //***********************************ProgramUpdateTypes***********************************
        public void Compare(Document doc) 
        {
            LibraryGetItems LibraryGetItems = new DBLibrary.LibraryGetItems();
            //ProgramUpdateTypes updateTypes = new ProgramUpdateTypes();

            List<Wall> ListWalls = LibraryGetItems.GetWalls(doc);
            Database_Excel excel = new Database_Excel();
            excel.ExcelRequest(@"W:\S\BIM\Z-LINKED EXCEL\WUHAN - WALLS.xlsx");
            List<ObjectWall> ListObjectWalls = excel.ExcelReadWallFile();

            foreach (Wall w in ListWalls)
            {
                Element e = w as Element;
                FamilyInstance familyInstance = e as FamilyInstance;
                Parameter TopConstraintParam = w.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                Parameter SOMIDParam = e.LookupParameter("SOM ID");

                string topConstraint = TopConstraintParam.AsValueString();
                string SOMID = LibraryGetItems.GetParameterValue(SOMIDParam);

                ObjectWall objectWall = ListObjectWalls.Find(x => x.Level == topConstraint && x.SOMID == SOMID);

                if (objectWall.Type != e.Name)
                {
                    changeFamilyType(doc, w, objectWall, ListWalls);
                }
            }
        }

        //***********************************GetTranslationExcelSheet***********************************
        public void changeFamilyType(Document doc, Wall wall, ObjectWall objectWall, List<Wall> ListWalls)
        {
            LibraryGetItems LibraryGetItems = new DBLibrary.LibraryGetItems();
            FamilySymbol familySymbol = LibraryGetItems.GetFamilySymbol(doc, objectWall.Type, BuiltInCategory.OST_Walls);


        }
    }
}
