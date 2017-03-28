using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.AutoCAD;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Windows;
using System.Runtime.InteropServices;
using acApp = Autodesk.AutoCAD.ApplicationServices.Application;
using kex;

namespace kex
{
    public class Kex1 : IExtensionApplication
    {
        private static Autodesk.AutoCAD.Windows.PaletteSet paletteSetKex = null;

        public void Initialize()
        {
            acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Cscript kex.dll is loaded.." + Environment.NewLine);
        }

        public void Terminate()
        {
            throw new NotImplementedException();
        }


        [CommandMethod("KEX")]
        public void StlEx()
        {
            if (paletteSetKex == null)
            {
                //MessageBox.Show("Palette 'KexControl1' don't exist. Try to create it...");
                try
                {
                    paletteSetKex = new PaletteSet("Kex - STL export");
                    KexControl1 myPal = new KexControl1();
                    paletteSetKex.Add("Kex - STL export", myPal);

                    //paletteSetKex.KeepFocus = true;
                    paletteSetKex.Visible = true;
                    //                      PaletteSetStyles.ShowAutoHideButton |
                    //                      PaletteSetStyles.ShowPropertiesMenu |
                    paletteSetKex.Style =
                      PaletteSetStyles.NameEditable |
                      PaletteSetStyles.ShowCloseButton;

                    //paletteSetKex.Location = new System.Drawing.Point(300, 400);
                    paletteSetKex.MinimumSize = new System.Drawing.Size(355, 700);
                    paletteSetKex.Size = new System.Drawing.Size(355, 700);
                    paletteSetKex.DockEnabled = DockSides.None;
                    paletteSetKex.Dock = DockSides.None;
                }
                catch
                {
                    MessageBox.Show("Failed to create the palette.");
                }
            }
            else
            {
                //MessageBox.Show("Palette 'KexControl1' exist already.");
                try
                {
                    paletteSetKex.Visible = true;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Original error: " + ex.Message);
                }
            }
        }

        [CommandMethod("RL")]
        public void RandomLayer()
        {
            DocumentCollection documentManager = acApp.DocumentManager;
            Document document = documentManager.MdiActiveDocument;
            Editor editor = document.Editor;
            Database database = document.Database;
            int counter = 1;
            bool displayErrorMessage = true;

            using (Transaction tr = database.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(database.LayerTableId, OpenMode.ForRead);
                PromptSelectionResult psr = editor.GetSelection();
                if(psr.Status == PromptStatus.OK)
                {
                    SelectionSet sset = psr.Value;
                    foreach(SelectedObject so in sset)
                    {
                        if(so != null)
                        {
                            try
                            {
                                Entity ent = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForWrite);
                                if (ent != null)
                                {
                                    string myNewLayer = "000_" + counter + "_" + ent.Layer.ToString();
                                    for (int i = 0; i < 10000; i++)
                                    {
                                        if (lt.Has(myNewLayer))
                                        {
                                            counter++;
                                            myNewLayer = "000_" + counter + "_" + ent.Layer.ToString();
                                        }
                                        else
                                        {
                                            lt.UpgradeOpen();
                                            LayerTableRecord ltr = new LayerTableRecord();
                                            ltr.Name = myNewLayer;
                                            lt.Add(ltr);
                                            tr.AddNewlyCreatedDBObject(ltr, true);
                                            break;
                                        }
                                    }
                                    try
                                    {
                                        ent.Layer = myNewLayer;
                                        counter++;
                                    }
                                    catch
                                    {
                                        if (displayErrorMessage)
                                        {
                                            displayErrorMessage = false;
                                            MessageBox.Show("Can't apply the new layer.");
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                if (displayErrorMessage)
                                {
                                    displayErrorMessage = false;
                                    MessageBox.Show("One or more objects are on locked or freezed layers.");
                                }
                            }
                        }
                    }

                }
                tr.Commit();
            }
        }

        //private void test()
        //{
        //    TypedValue[] tv = new TypedValue[4] {
        //                new TypedValue(67, 0), //Model space
        //                new TypedValue((int)DxfCode.Visibility, 0),
        //                new TypedValue((int)DxfCode.LayerName, layer.Name),
        //                new TypedValue((int)DxfCode.Start, "INSERT,3DSOLID"),
        //                };
        //    SelectionFilter sf = new SelectionFilter(tv);
        //    PromptSelectionResult res = documentManager.MdiActiveDocument.Editor.SelectAll(sf);
        //    SelectionSet copySset = res.Value;

        //    int c = 0;
        //    using (Transaction tr = database.TransactionManager.StartTransaction())
        //    {
        //        foreach (SelectedObject so in copySset)
        //        {

        //            Entity ent = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForRead);
        //            if (ent.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.Solid3d")
        //            {
        //                using (var brep = new Brep(ent))
        //                {
        //                    c += brep.Vertices.Count();
        //                }
        //            }
        //            else if (ent.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.BlockReference")
        //            {
        //                MessageBox.Show("Block found: (" + ent.GetType().ToString() + ").");
        //                using (DBObjectCollection dbObjCol = new DBObjectCollection())
        //                {
        //                    try
        //                    {
        //                        BlockReference bref = ent as BlockReference;
        //                        bref.Explode(dbObjCol);

        //                        try
        //                        {
        //                            foreach (DBObject dbObj in dbObjCol)
        //                            {
        //                                Entity acEnt = dbObj as Entity;
        //                                BlockTableRecord acCurSpaceBlkTblRec = tr.GetObject(database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
        //                                acCurSpaceBlkTblRec.AppendEntity(acEnt);
        //                                tr.AddNewlyCreatedDBObject(dbObj, true);


        //                                using (var brep = new Brep(acEnt))
        //                                {
        //                                    c += brep.Vertices.Count();
        //                                }
        //                            }
        //                        }
        //                        catch (System.Exception ex)
        //                        {
        //                            MessageBox.Show("Can't count exploded object vertices: (" + bref.Name.ToString() + ")." + "  | Original error: " + ex.Message);
        //                        }
        //                    }
        //                    catch (System.Exception ex)
        //                    {
        //                        MessageBox.Show("Can't explode the block: (" + ent.BlockName.ToString() + ")." + "  | Original error: " + ex.Message);
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("Undefined object found: (" + ent.GetType().ToString() + ").");
        //            }
        //        }
        //        tr.Commit();
        //    }
        //    MessageBox.Show("Vertices: (" + c + ").");

        //}
    }
}
