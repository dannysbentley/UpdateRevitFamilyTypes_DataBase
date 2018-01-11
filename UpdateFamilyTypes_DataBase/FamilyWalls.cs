using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateFamilyTypes_DataBase
{
    partial class FamilyWall
    {
        public Element element { get; set; }
        public Family family { get; set; }
        public string SOMID { get; set; }
        public string TypeName { get; set; }
    }
}
