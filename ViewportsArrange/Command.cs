#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace ViewportsArrange
{
    /// <summary>
    /// Implements the Revit add-in interface IExternalCommand
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class AlignHorizontal : IExternalCommand
    {
        public static String global_message;

        public virtual Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Util align = new Util("Horizontal", uidoc, doc);
            align.Align();

            return Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class AlignVertical : IExternalCommand
    {
        public static String global_message;

        public virtual Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Util align = new Util("Vertical", uidoc, doc);
            align.Align();

            return Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class AlignTop : IExternalCommand
    {
        public static String global_message;

        public virtual Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Util align = new Util("Top", uidoc, doc);
            align.Align();

            return Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class AlignBottom : IExternalCommand
    {
        public static String global_message;

        public virtual Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Util align = new Util("Bottom", uidoc, doc);
            align.Align();

            return Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class AlignLeft: IExternalCommand
    {
        public static String global_message;

        public virtual Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Util align = new Util("Left", uidoc, doc);
            align.Align();

            return Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class AlignRight : IExternalCommand
    {
        public static String global_message;

        public virtual Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Util align = new Util("Right", uidoc, doc);
            align.Align();

            return Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class DistributeHorizontal : IExternalCommand
    {
        public static String global_message;

        public virtual Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Util distribute = new Util("Horizontal", uidoc, doc);
            distribute.Distribute();

            return Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class DistributeVertical : IExternalCommand
    {
        public static String global_message;

        public virtual Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Util distribute = new Util("Vertical", uidoc, doc);
            distribute.Distribute();

            return Result.Succeeded;
        }
    }
    public class Util
    {
        private UIDocument uidoc;
        private Document doc;
        private ViewPortSelectionFilter filter;
        private IList<Reference> viewports;
        private string type;

        /// <summary>
        /// Initializes a new instance of the <see cref="Util"/> class.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="uidoc">The uidoc.</param>
        /// <param name="doc">The document.</param>
        public Util(string type, UIDocument uidoc, Document doc)
        {
            this.uidoc = uidoc;
            this.doc = doc;
            this.filter = new ViewPortSelectionFilter(doc);
            this.viewports = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select multiple viewports");
            this.type = type;
        }
        /// <summary>
        /// Main align method of the class.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="viewports">The viewports.</param>
        public void Align()
        {
            if (viewports.Count > 1)
            {
                Align(doc, viewports, type);
            }	
        }
        /// <summary>
        /// Main distribute method of the class.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="viewports">The viewports.</param>
        public void Distribute()
        {
            if (viewports.Count > 1)
            {
                Distribute(doc, viewports, type);
            }
        }
        /// <summary>
        /// Distributes the selected viewports.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="viewports">The viewports.</param>
        /// <param name="direction">The direction.</param>
		private void Distribute(Document doc, IList<Reference> viewports, string direction)
		{		
			Viewport [] viewportArray = null;			
			double [] centers = null;			
			
			SetValues(doc, viewports, direction, ref viewportArray, ref centers);		
			
			double distance = (centers[centers.Length - 1] - centers[0])/(centers.Length - 1);
			double delta = 0;
			XYZ translation = null;
			
			using(Transaction t = new Transaction(doc, "Arrange Viewports"))
			{				
				t.Start();
				for (int i = 0; i < centers.Length - 1; i++)
				{
					delta = (i*distance - (centers[i]-centers[0]));
					switch(direction)
					{
						case "Vertical": case "Top": case "Bottom":							
							translation = new XYZ(0, delta, 0);
							break;
						case "Horizontal": case "Left": case "Right":						
							translation = new XYZ(delta, 0, 0);
							break;							
					}
					ElementTransformUtils.MoveElement(doc, viewportArray[i].Id, translation);
				}
				t.Commit();
			}
			return;
		}
        /// <summary>
        /// Aligns the selected viewports.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="viewports">The viewports.</param>
        /// <param name="direction">The direction.</param>
		private void Align(Document doc, IList<Reference> viewports, string direction)
		{		
			Viewport [] viewportArray = null;			
			double [] centers = null;
			
			SetValues(doc, viewports, direction, ref viewportArray, ref centers);
			
			double distance = (centers[centers.Length - 1] + centers[0])/2;
			double delta = 0;
			
			XYZ translation = null;
			
			using(Transaction t = new Transaction(doc, "Arrange Viewports"))
			{				
				t.Start();
				for (int i = 0; i < centers.Length; i++)
				{
					delta = distance-centers[i];
					if(direction == "Vertical" || direction == "Top" || direction == "Bottom")
					{
						translation = new XYZ(0, delta, 0);						
					}
					else 
					{						
						translation = new XYZ(delta, 0, 0);	
					}
					ElementTransformUtils.MoveElement(doc, viewportArray[i].Id, translation);
				}
				t.Commit();
			}
			return;
		}
        /// <summary>
        /// Sets the values.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="viewports">The viewports.</param>
        /// <param name="direction">The direction.</param>
        /// <param name="viewportArray">The viewport array.</param>
        /// <param name="centers">The centers.</param>
		private void SetValues(Document doc, IList<Reference> viewports, string direction, ref Viewport [] viewportArray, ref double [] centers)
		{			
			switch(direction)
			{
				case "Vertical":					
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxCenter().Y).ToArray();
					centers = viewportArray.Select(x => x.GetBoxCenter().Y).ToArray();
					break;
				case "Horizontal":					
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxCenter().X).ToArray();
					centers = viewportArray.Select(x => x.GetBoxCenter().X).ToArray();
					break;
				case "Top":
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxOutline().MinimumPoint.Y).ToArray();
					centers = viewportArray.Select(x => x.GetBoxOutline().MinimumPoint.Y).ToArray();					
					break;
				case "Bottom":
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxOutline().MaximumPoint.Y).ToArray();
					centers = viewportArray.Select(x => x.GetBoxOutline().MaximumPoint.Y).ToArray();					
					break;
				case "Left":
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxOutline().MinimumPoint.X).ToArray();
					centers = viewportArray.Select(x => x.GetBoxOutline().MinimumPoint.X).ToArray();					
					break;
				case "Right":
					viewportArray = viewports.Select(i => doc.GetElement(i) as Viewport).OrderBy(o => o.GetBoxOutline().MaximumPoint.X).ToArray();
					centers = viewportArray.Select(x => x.GetBoxOutline().MaximumPoint.X).ToArray();					
					break;
			}		
		}
        /// <summary>
        /// Fillters only viewport elements
        /// </summary>
        /// <seealso cref="Autodesk.Revit.UI.Selection.ISelectionFilter" />
        internal class ViewPortSelectionFilter : ISelectionFilter
        {
            Document doc = null;
            public ViewPortSelectionFilter(Document document)
            {
                doc = document;
            }

            public bool AllowElement(Element element)
            {
                if (element.Category.Name == "Viewports")
                {
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return true;
            }
        }
    }
}
