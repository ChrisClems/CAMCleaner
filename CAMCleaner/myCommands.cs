// (C) Copyright 2024 by  
//

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(CAMCleaner.MyCommands))]

namespace CAMCleaner
{
    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class MyCommands
    {
        /// <summary>
        /// Adjust the normals of selected lightweight polylines (LWPOLYLINE)
        /// in the current drawing to be oriented along the positive Z-axis (0, 0, 1).
        /// Any polyline normals that are not already oriented in this manner
        /// will be adjusted, and their vertex coordinates will be converted
        /// to 2D points by discarding the Z-coordinate component and flipping
        /// the associated bulge values.
        /// </summary>
        /// <remarks>
        /// This command uses a selection filter to prompt the user to select
        /// LWPOLYLINE entities. Only polylines that do not already have their
        /// normals oriented along the positive Z-axis will be modified. The
        /// command outputs the number of polylines that were adjusted.
        /// </remarks>
        /// <example>
        /// AutoCAD command: FlattenPolyNormals
        /// </example>
        [CommandMethod("CNChris", "FlattenPolyNormals", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void FlattenPolyNormals()
        {
            Vector3d normalZ2d = new Vector3d(0, 0, 1);
            Document doc = Application.DocumentManager.MdiActiveDocument;;
            Database acCurDb = doc.Database;
            var ed = doc.Editor;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                int fixedPolyCount = 0;
                
                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                };
                SelectionFilter filter = new SelectionFilter(filterList);
                
                PromptSelectionOptions opts = new PromptSelectionOptions
                {
                    MessageForAdding = "\nSelect polylines: ",
                    AllowDuplicates = false
                };
                
                PromptSelectionResult acSSPrompt = ed.GetSelection(opts, filter);

                if (acSSPrompt.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nNo objects selected ");
                    return;
                }
                
                if (acSSPrompt.Status == PromptStatus.OK && acSSPrompt.Value.Count > 0)
                {
                    var acAllObjs = acSSPrompt.Value.GetObjectIds();
                    foreach (var acObj in acAllObjs)
                    {
                        if (acObj.ObjectClass.Name == "AcDbPolyline")
                        {
                            var polyline = acTrans.GetObject(acObj, OpenMode.ForWrite) as Polyline;
                            if (polyline.Normal != normalZ2d)
                            {
                                fixedPolyCount++;
                                for (int i = 0; i < polyline.NumberOfVertices; i++)
                                {
                                    //var acPlArc = acPlLwObj.GetArcSegmentAt(i);
                                    var acPl3DPoint = polyline.GetPoint3dAt(i);
                                    var acPl2DPointNew = new Point2d(acPl3DPoint.X, acPl3DPoint.Y);
                                    polyline.SetPointAt(i, acPl2DPointNew);
                                    polyline.SetBulgeAt(i, -polyline.GetBulgeAt(i));
                                }
                                polyline.Normal = normalZ2d;
                            }
                        }
                    }
                    acTrans.Commit();
                }
                ed.WriteMessage($@"Fixed {fixedPolyCount} inverted polyline normals.");
            }
        }

        /// <summary>
        /// Executes a combination of AUDIT and PURGE commands on the current drawing.
        /// This method first performs an AUDIT with the "Yes" (Y) option to repair any
        /// errors found in the drawing’s database. Following this, it performs two
        /// PURGE commands: one to remove registered applications (regapps) and one
        /// to remove all other purgeable objects. Both PURGE commands are executed
        /// with the "No" (N) option, meaning no confirmation is sought from the user
        /// for each item to be purged.
        /// </summary>
        /// <remarks>
        /// The method provides a macro for the basic drawing repair operations suggested in the official docs:
        /// https://www.autodesk.com/support/technical/article/caas/sfdcarticles/sfdcarticles/Optimizing-the-AutoCAD-drawing-file-Purge-Audit-Recover.html
        /// </remarks>
        [CommandMethod("CNChris", "AuditPurge", CommandFlags.Modal)]
        public void AuditPurge()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            acDoc.SendStringToExecute("AUDIT Y ", true, false, true);
            acDoc.SendStringToExecute("-PURGE A *\nN ", true, false, true);
            acDoc.SendStringToExecute("-PURGE R *\nN ", true, false, true);
        }

        /// <summary>
        /// Normalizes the text styles for all TEXT and MTEXT entities in the current drawing.
        /// This method iterates through all text entities and applies the current annotation style
        /// ensuring uniformity and consistency across the drawing.
        /// </summary>
        /// <remarks>
        /// This command uses a selection filter to automatically select all TEXT and MTEXT entities
        /// within the current drawing. If no such entities are found, a message will be displayed to
        /// the user indicating that no text entities were found. The command adjusts the styles of
        /// these text entities to match the default or specified settings.
        /// </remarks>
        /// <summary>
        /// AutoCAD command: NormalizeTextStyles
        /// </summary>
        [CommandMethod("CNChris", "NormalizeTextStyles", CommandFlags.Modal)]
        public void NormalizeTextStyles()
        {
            //TODO: Rewrite to not require pre-selecting current default style
            Document doc = Application.DocumentManager.MdiActiveDocument;
            var database = doc.Database;
            var ed = doc.Editor;
                
            using (Transaction tr = database.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
            
                var filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<or"),
                    new TypedValue((int)DxfCode.Start, "TEXT"),
                    new TypedValue((int)DxfCode.Start, "MTEXT"),
                    new TypedValue((int)DxfCode.Operator, "or>")
                };
            
                SelectionFilter filter = new SelectionFilter(filterList);
                PromptSelectionResult psr = ed.SelectAll(filter);
            
                if (psr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nNo text entities found");
                    return;
                }

                SelectionSet ss = psr.Value;

                ObjectId[] objectIdArray = ss.GetObjectIds();
                ObjectIdCollection objectIdCollection = new ObjectIdCollection(objectIdArray);

                ObjectIdCollection markText = new ObjectIdCollection(); 
                foreach (ObjectId id in objectIdCollection)
                {
                    if (id.ObjectClass.DxfName == "TEXT")
                    {
                        DBText text = (DBText)tr.GetObject(id, OpenMode.ForWrite);
                        text.TextStyleId = database.Textstyle;
                    }
                    else if (id.ObjectClass.DxfName == "MTEXT")
                    {
                        var text = (MText)tr.GetObject(id, OpenMode.ForWrite);
                        text.TextStyleId = database.Textstyle;
                    }
                }
                tr.Commit();
                // var textStyle = (TextStyleTableRecord)database.Textstyle.GetObject(OpenMode.ForRead);
                // var textStyleName = textStyle.Name;
                // ed.WriteMessage($"All text entities changed to style {textStyleName}");
            }
        }
    }
}