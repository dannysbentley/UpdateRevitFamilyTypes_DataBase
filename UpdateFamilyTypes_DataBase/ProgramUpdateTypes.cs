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
        //***********************************Compare***********************************
        public void Compare(Document doc) 
        {
            LibraryGetItems LibraryGetItems = new DBLibrary.LibraryGetItems();
            //ProgramUpdateTypes updateTypes = new ProgramUpdateTypes();

            List<Wall> ListWalls = LibraryGetItems.GetWalls(doc);
            List<FamilyInstance> ListFamilyInstance = GetFamilyInstance(doc, BuiltInCategory.OST_Walls);
            Database_Excel excel = new Database_Excel();
            excel.ExcelRequest(@"W:\S\BIM\Z-LINKED EXCEL\WUHAN - WALLS.xlsx");
            List<ObjectWall> ListObjectWalls = excel.ExcelReadWallFile();

            foreach (Wall w in ListWalls)
            {
                Element e = w as Element;
                FamilyInstance familyInstance = ListFamilyInstance.Find(x => x.Symbol.Name == w.Name);

                Parameter TopConstraintParam = w.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                Parameter SOMIDParam = e.LookupParameter("SOM ID");

                string topConstraint = TopConstraintParam.AsValueString();
                string SOMID = LibraryGetItems.GetParameterValue(SOMIDParam);

                ObjectWall objectWall = ListObjectWalls.Find(x => x.Level == topConstraint && x.SOMID == SOMID);

                if (objectWall.Type != e.Name)
                {
                    changeFamilyType(doc, familyInstance, objectWall);
                }
            }
        }

        //***********************************changeFamilyType***********************************
        public void changeFamilyType(Document doc, FamilyInstance familyInstance, ObjectWall objectWall)
        {
            LibraryGetItems LibraryGetItems = new DBLibrary.LibraryGetItems();
            FamilySymbol familySymbol = LibraryGetItems.GetFamilySymbol(doc, objectWall.Type, BuiltInCategory.OST_Walls);

            // Transaction to change the element type 
            Transaction trans = new Transaction(doc, "Edit Type");
            trans.Start();
            try
            {
                familyInstance.Symbol = familySymbol;
            }
            catch { }
            trans.Commit();
        }

        public List<FamilyInstance> GetFamilyInstance(Document doc, BuiltInCategory category)
        {
            List<FamilyInstance> ListFamilyInstance = new List<FamilyInstance>();
            List<Element> elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsElementType().ToList<Element>();

            foreach (Element e in elements)
            {
                FamilyInstance familyInstance = e as FamilyInstance;
                ListFamilyInstance.Add(familyInstance);
            }
            return ListFamilyInstance;
        }
    }
}
