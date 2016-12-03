// Copyright notice

// Archilizer Family Interface
// to be used in conjunction with Revit Family Editor
// Femily Editor Interface is ease of use type of plug-in -
// it main purpose is to make the process of creating and editting
// of Revit families smoother and more pleasent (less time consuming too)

#region Namespaces
using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Windows.Media.Imaging;
using System.Linq;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System.Windows.Forms;
using System.Diagnostics;
#endregion

namespace ViewportsArrange
{
    /// <summary>
    /// Implements the Revit add-in interface IExternalApplication
    /// </summary>
    class App : IExternalApplication
    {
        // class instance
        internal static App thisApp = null;
        public const string Message = "Aligns and distributes multiple viewports on a sheet.";
        /// <summary>
        /// Get absolute path to this assembly
        /// </summary>
        static string path = Assembly.GetExecutingAssembly().Location;
        static string contentPath = Path.GetDirectoryName(Path.GetDirectoryName(path)) + "/";
        //static string helpFile = "file:///C:/ProgramData/Autodesk/ApplicationPlugins/ViewportsArrange.bundle/Content/Family%20Editor%20Interface%20_%20Revit%20_%20Autodesk%20App%20Store.html";
        static string largeIcon = contentPath + "modeless32.png";
        static string smallIcon = contentPath + "modeless16.png";
        SplitButton vpArrange;
        #region Ribbon
        /// <summary>
        /// Use embedded image to load as an icon for the ribbon
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        static private BitmapSource GetEmbeddedImage(string name)
        {
            try
            {
                Assembly a = Assembly.GetExecutingAssembly();
                Stream s = a.GetManifestResourceStream(name);
                return BitmapFrame.Create(s);
            }
            catch
            {
                return null;
            }
        }
        /// <summary>
        /// Add ribbon panel 
        /// </summary>
        /// <param name="a"></param>
        private void AddRibbonPanel(UIControlledApplication a)
        {
            List<RibbonPanel> panels = a.GetRibbonPanels();
            Autodesk.Revit.UI.RibbonPanel rvtRibbonPanel = null;
            if (panels.FirstOrDefault(x => x.Name.Equals("Archilizer", StringComparison.OrdinalIgnoreCase)) == null)
            {
                rvtRibbonPanel = a.CreateRibbonPanel("Archilizer");
            }
            else
            {
                rvtRibbonPanel = panels.FirstOrDefault(x => x.Name.Equals("Archilizer", StringComparison.OrdinalIgnoreCase)) as RibbonPanel;
            }

            PushButtonData align_v = new PushButtonData("align vertical", "Align Vertical", path,
                "ViewportsArrange.AlignVertical");
            PushButtonData align_h = new PushButtonData("align horizontal", "Align Horizontal", path,
                "ViewportsArrange.AlignHorizontal");
            PushButtonData align_t = new PushButtonData("align top", "Align Top", path,
                "ViewportsArrange.AlignTop");
            PushButtonData align_b = new PushButtonData("align bottom", "Align Bottom", path,
                "ViewportsArrange.AlignBottom");
            PushButtonData align_l = new PushButtonData("align left", "Align Left", path,
                "ViewportsArrange.AlignLeft");
            PushButtonData align_r = new PushButtonData("align right", "Align Right", path,
                "ViewportsArrange.AlignRight");
            PushButtonData distribute_v = new PushButtonData("`distribute vertical", "Distribute Vertical", path,
                "ViewportsArrange.DistributeVertical");
            PushButtonData distribute_h = new PushButtonData("`distribute horizontal", "Distribute Horizontal", path,
                "ViewportsArrange.DistributeHorizontal");
            align_v.LargeImage = new BitmapImage(new Uri(@smallIcon));
            align_h.LargeImage = new BitmapImage(new Uri(@smallIcon));
            align_t.LargeImage = new BitmapImage(new Uri(@smallIcon));
            align_b.LargeImage = new BitmapImage(new Uri(@smallIcon));
            align_l.LargeImage = new BitmapImage(new Uri(@smallIcon));
            align_r.LargeImage = new BitmapImage(new Uri(@smallIcon));
            distribute_v.LargeImage = new BitmapImage(new Uri(@smallIcon));
            distribute_h.LargeImage = new BitmapImage(new Uri(@smallIcon));

            /*
            SplitButtonData vpArrangeData = new SplitButtonData("ViewportsArrange", "Viewports" + Environment.NewLine + "Arrange" + Environment.NewLine);
            vpArrange = rvtRibbonPanel.AddItem(vpArrangeData) as SplitButton;
            vpArrange.AddPushButton(align_h);
            vpArrange.AddPushButton(align_v);
            */

            PulldownButtonData data = new PulldownButtonData("ViewportsArrange", "Viewports Arrange");
            PulldownButton viewportsArrange = rvtRibbonPanel.AddItem(data) as PulldownButton;

            viewportsArrange.AddPushButton(align_v);
            viewportsArrange.AddPushButton(align_h);
            viewportsArrange.AddPushButton(align_t);
            viewportsArrange.AddPushButton(align_b);
            viewportsArrange.AddPushButton(align_l);
            viewportsArrange.AddPushButton(align_r);
            viewportsArrange.AddPushButton(distribute_v);
            viewportsArrange.AddPushButton(distribute_h);

            BitmapSource img32 = new BitmapImage(new Uri(@largeIcon));
            BitmapSource img16 = new BitmapImage(new Uri(@smallIcon));

            //ContextualHelp ch = new ContextualHelp(ContextualHelpType.Url, @helpFile);
            
            viewportsArrange.Image = img16;
            viewportsArrange.LargeImage = img32;
            viewportsArrange.ToolTip = Message;
            //familyEI.SetContextualHelp(ch);
        }
        #endregion

        public Result OnStartup(UIControlledApplication a)
        {
            ControlledApplication c_app = a.ControlledApplication;
            AddRibbonPanel(a);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
