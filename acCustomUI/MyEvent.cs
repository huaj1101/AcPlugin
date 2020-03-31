using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(AutoCAD_CSharp_plug_in_acCustomUI.MyEvent))]

namespace AutoCAD_CSharp_plug_in_acCustomUI
{
    public class MyEvent
    {
        static Autodesk.AutoCAD.EditorInput.Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

        [CommandMethod("Addselectchang")]
        public static void AddDocEvent()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            acDoc.ImpliedSelectionChanged += new EventHandler(doc_ImpliedSelectionChanged);
        }

        [CommandMethod("Removeselectchang")]
        public static void RemoveDocEvent()
        {
            // Get the current document
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            acDoc.ImpliedSelectionChanged -= new EventHandler(doc_ImpliedSelectionChanged);
        }

        public static void doc_ImpliedSelectionChanged(object sender, EventArgs e)
        {
            PromptSelectionResult pkf = ed.SelectImplied();
            if (pkf.Status != PromptStatus.OK) return;
            ObjectId[] objIds = pkf.Value.GetObjectIds();
            String oids = "";

            foreach (ObjectId objId in objIds)
            {
                oids += "\n " + objId.ToString();
            }

            MyDockBarLeft.showSelectedObjectsInfo(oids);
            //Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("\n选择的对像ObjectId为：" +
            //                                                                 oids + "\n共选择了对像个数是:" + objIds.Length.ToString());
        }
    }
}
