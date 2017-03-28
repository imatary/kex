using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.BoundaryRepresentation;
using acApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Diagnostics;
using System.Xml;
using System.Threading;

namespace kex
{
    public partial class KexControl1 : UserControl
    {
        public string var_login = "";
        public double var_facetres = 10;
        public double progress = 0;
        public AcadDocument thisDrawing = (AcadDocument)acApp.DocumentManager.MdiActiveDocument.GetAcadDocument();
        public string dwg_filename = "";
        public string customProperty = "";
        public int currentLevel = 0;
        public string separator = "_";
        public bool running = false;
        public List<string> ignoreBlocks = new List<string>();
        public int blockCounterNestedBlocks = 0;
        //Open palette and load settings
        public KexControl1()
        {
            InitializeComponent();

            //Stop Button
            button10.Visible = false;

            var_login = thisDrawing.GetVariable("LOGINNAME");
            var_facetres = thisDrawing.GetVariable("FACETRES");
            textBox2.Text = var_login.ToString();
            textBox3.Text = Properties.Settings.Default.DEFAULT_FACETRES.ToString();
            textBox8.Text = Properties.Settings.Default.STL_PATH.ToString();
            textBox9.Text = Properties.Settings.Default.GROUP_PREFIX.ToString();
            checkBox1.Checked = Properties.Settings.Default.RESET_UCS;
            checkBox2.Checked = Properties.Settings.Default.MOVE_XYZ;
            checkBox3.Checked = Properties.Settings.Default.EXPLODE;
            checkBox5.Checked = Properties.Settings.Default.USE_NEW_DRAWING;
            textBox1.Text = Properties.Settings.Default.EXCEPTION_LAYERS;
            checkBox4.Checked = Properties.Settings.Default.EXCEPTION_ON;

            if (Properties.Settings.Default.LAYER_PER_BLOCK)
            {
                radioButton1.Checked = true;
            }
            else
            {
                radioButton2.Checked = true;
            }

            CheckExplodeBox();
            textBox10.Text = "";
            progressBar1.Value = progressBar1.Minimum;
            checkBox6.Checked = Properties.Settings.Default.USE_C4D_SCRIPT;
            checkBox7.Checked = Properties.Settings.Default.USE_CUSTOM_PROPERTIES;
            textBox4.Text = Properties.Settings.Default.PY_PATH;
            textBox5.Text = Properties.Settings.Default.PROPERTY_NAME;

            if(Properties.Settings.Default.USE_C4D_SCRIPT)
            {
                label5.Enabled = true;
                button8.Enabled = true;
                textBox4.Enabled = true;
            }
            else
            {
                label5.Enabled = false;
                button8.Enabled = false;
                textBox4.Enabled = false;
            }

            if (Properties.Settings.Default.USE_CUSTOM_PROPERTIES)
            {
                label6.Enabled = true;
                textBox5.Enabled = true;
                label7.Enabled = true;
                label8.Enabled = true;
            }
            else
            {
                label6.Enabled = false;
                textBox5.Enabled = false;
                label7.Enabled = false;
                label8.Enabled = false;
            }
            label13.Text = "Status:";

            //Milliseconds for timeout
            comboBox1.Items.Add("100");
            comboBox1.Items.Add("200");
            comboBox1.Items.Add("300");
            comboBox1.Items.Add("400");
            comboBox1.Items.Add("500");
            comboBox1.Items.Add("600");
            comboBox1.Items.Add("700");
            comboBox1.Items.Add("800");
            comboBox1.Items.Add("900");
            comboBox1.SelectedIndex = Properties.Settings.Default.TIMEOUT_SELECTEDITEM;
            checkBox8.Checked = Properties.Settings.Default.USE_TIMEOUT;
            if(Properties.Settings.Default.USE_TIMEOUT)
            {
                label14.Enabled = true;
                comboBox1.Enabled = true;
            }
            else
            {
                label14.Enabled = false;
                comboBox1.Enabled = false;
            }

            checkBox9.Checked = Properties.Settings.Default.NESTED_BLOCKS;
            if (Properties.Settings.Default.EXPLODE)
            {
                checkBox9.Enabled = true;
            }
            else
            {
                checkBox9.Enabled = false;
            }
            checkBox10.Checked = Properties.Settings.Default.CLOSE_NEW_DWG;
            if (Properties.Settings.Default.USE_NEW_DRAWING)
            {
                checkBox10.Enabled = true;
                label18.Enabled = true;
            }
            else
            {
                checkBox10.Enabled = false;
                label18.Enabled = false;
            }
            checkBox11.Checked = Properties.Settings.Default.SOLID_OUTSIDE_BLOCKS;

            if(Properties.Settings.Default.STLOUT)
            {
                radioButton4.Checked = true;
            }
            else
            {
                radioButton3.Checked = true;
            }

        }

        //Enable or Disable Radio Buttons (Layer/Block or Objects/Block)
        public void CheckExplodeBox()
        {
            if (checkBox3.Checked)
            {
                checkBox9.Enabled = true;
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                checkBox4.Enabled = true;
                if (checkBox4.Checked)
                {
                    label2.Enabled = true;
                    textBox1.Enabled = true;
                }
                else
                {
                    label2.Enabled = false;
                    textBox1.Enabled = false;
                }
            }
            else
            {
                checkBox9.Enabled = false;
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                checkBox4.Enabled = false;
                label2.Enabled = false;
                textBox1.Enabled = false;
            }
        }

        //Save settings button
        private void button5_Click(object sender, EventArgs e)
        {
            SaveSettings();
            EditPythomScript();
        }

        //Save settings
        private void SaveSettings()
        {
            Properties.Settings.Default.DEFAULT_FACETRES = textBox3.Text;
            Properties.Settings.Default.STL_PATH = textBox8.Text;
            Properties.Settings.Default.GROUP_PREFIX = textBox9.Text;
            Properties.Settings.Default.RESET_UCS = checkBox1.Checked;
            Properties.Settings.Default.MOVE_XYZ = checkBox2.Checked;
            Properties.Settings.Default.EXPLODE = checkBox3.Checked;
            Properties.Settings.Default.EXCEPTION_ON = checkBox4.Checked;
            Properties.Settings.Default.EXCEPTION_LAYERS = textBox1.Text;
            Properties.Settings.Default.USE_NEW_DRAWING = checkBox5.Checked;
            Properties.Settings.Default.USE_C4D_SCRIPT = checkBox6.Checked;
            Properties.Settings.Default.PY_PATH = textBox4.Text;
            Properties.Settings.Default.USE_CUSTOM_PROPERTIES = checkBox7.Checked;
            Properties.Settings.Default.PROPERTY_NAME = textBox5.Text;
            Properties.Settings.Default.USE_TIMEOUT = checkBox8.Checked;
            Properties.Settings.Default.TIMEOUT_SELECTEDITEM = comboBox1.SelectedIndex;
            Properties.Settings.Default.NESTED_BLOCKS = checkBox9.Checked;
            Properties.Settings.Default.CLOSE_NEW_DWG = checkBox10.Checked;
            Properties.Settings.Default.SOLID_OUTSIDE_BLOCKS = checkBox11.Checked;
            Properties.Settings.Default.STLOUT = radioButton4.Checked;
            Properties.Settings.Default.Save();
        }

        //Change .py script variables
        void EditPythomScript()
        {
            //Change .py script variables
            if (checkBox6.Checked)
            {
                if (File.Exists(textBox4.Text))
                {
                    int lineNumber = 0;
                    string oldLineText = "";
                    try
                    {
                        foreach (string line in File.ReadLines(@textBox4.Text))
                        {
                            if (line.Contains("exportpath ="))
                            {
                                //MessageBox.Show(line);
                                oldLineText = line;
                                break;
                            }
                            lineNumber++;
                        }
                        if (lineNumber > 0)
                        {
                            string newLineText = "\texportpath = \"" + textBox8.Text.Replace("\\", "/") + "\"";
                            //MessageBox.Show("New text for this line: " + newLineText);
                            if (newLineText != oldLineText)
                            {
                                lineChanger(newLineText, @textBox4.Text, lineNumber);
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("Can't edit .py script: " + textBox4.Text);
                    }
                }
                else
                {
                    MessageBox.Show(".py script for Cinema4D not found: " + textBox4.Text);
                }
            }
        }

        //Replace a line from a file with a new line
        static void lineChanger(string newText, string fileName, int line_to_edit)
        {
            string[] arrLine = File.ReadAllLines(fileName);
            arrLine[line_to_edit] = newText;
            File.WriteAllLines(fileName, arrLine);
            //MessageBox.Show("Saved: " + fileName);
        }

        //OpdenDialog for STL Path
        private void button2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            if (textBox8.Text != "")
            {
                folderDialog.SelectedPath = textBox8.Text;
            }
            else
            {
                folderDialog.SelectedPath = "c:\\";
            }

            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox8.Text = folderDialog.SelectedPath;

                    //Change .py script variables
                    EditPythomScript();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Error: Could not read folder from disk. Original error: " + ex.Message);
                }
            }
        }

        //Explode blocks checkbox
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            CheckExplodeBox();
        }

        //Exception layers checkbox
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                label2.Enabled = true;
                textBox1.Enabled = true;
            }
            else
            {
                label2.Enabled = false;
                textBox1.Enabled = false;
            }
        }

        //Quality / Facetres min. value
        private void button3_Click(object sender, EventArgs e)
        {
            textBox3.Text = "0.01";
        }

        //Quality / Facetres max. value
        private void button4_Click(object sender, EventArgs e)
        {
            textBox3.Text = "10";
        }

        //PY Script checkbox
        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox6.Checked)
            {
                label5.Enabled = true;
                button8.Enabled = true;
                textBox4.Enabled = true;
            }
            else
            {
                label5.Enabled = false;
                button8.Enabled = false;
                textBox4.Enabled = false;
            }

        }

        //Use custom properties checkbox
        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                label6.Enabled = true;
                textBox5.Enabled = true;
                label7.Enabled = true;
                label8.Enabled = true;
            }
            else
            {
                label6.Enabled = false;
                textBox5.Enabled = false;
                label7.Enabled = false;
                label8.Enabled = false;
            }
        }

        //Select PY path button
        private void button8_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            if (textBox8.Text != "")
            {
                openFileDialog1.InitialDirectory = textBox4.Text;
            }
            else
            {
                openFileDialog1.InitialDirectory = "c:\\" + "users\\" + var_login + "\\";
            }
            openFileDialog1.Filter = "py files (*.py)|*.py|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox4.Text = openFileDialog1.FileName;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        //Timeout chackbox
        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox8.Checked)
            {
                label14.Enabled = true;
                comboBox1.Enabled = true;
            }
            else
            {
                label14.Enabled = false;
                comboBox1.Enabled = false;
            }
        }

        //Use new drawing checkbox
        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                checkBox10.Enabled = true;
                label18.Enabled = true;
            }
            else
            {
                checkBox10.Enabled = false;
                label18.Enabled = false;
            }
        }

        //Clear Results Button
        private void button6_Click(object sender, EventArgs e)
        {
            textBox10.Text = "";
            label12.Text = "...";
        }

        //Open Folder Button
        private void button7_Click(object sender, EventArgs e)
        {
            Process.Start(@Properties.Settings.Default.STL_PATH.ToString());
        }

        //Save Facetres
        private void SaveFacetres()
        {
            var_facetres = thisDrawing.GetVariable("FACETRES");
            try
            {
                double exportQuality = Convert.ToDouble(textBox3.Text);
                thisDrawing.SetVariable("FACETRES", exportQuality);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Invalid value for FACETRES. Can't change export quality. Original error: " + ex.Message);
            }
        }

        //Restore Facetres
        private void RestoreFacetres()
        {
            thisDrawing.SetVariable("FACETRES", var_facetres);
        }

        //HELPER: Restore Filedia
        private void button9_Click(object sender, EventArgs e)
        {
            thisDrawing.SetVariable("FILEDIA", 1);
        }

        //Delete special characters
        public static string RemoveSpecialCharacters(string str)
        {
            str = str.Replace(" ", "-");
            str = str.Replace("_", "-");
            str = str.Replace("+", "-");
            str = str.Replace("&", "-");
            str = str.Replace("Ö", "OE");
            str = str.Replace("Ä", "AE");
            str = str.Replace("Ü", "UE");
            str = str.Replace("É", "E");
            str = str.Replace("À", "A");
            str = str.Replace("È", "E");
            str = str.Replace("°", "GRD");

            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        //Progressbar update
        private int UpdateProgressBar(double progressStep)
        {
            //Update progressBar
            progress += progressStep;
            if (progress > progressBar1.Maximum)
            {
                progress = progressBar1.Maximum;
            }
            return (int)Math.Round(progress, 0);
        }

        //Stop export
        private void button10_Click(object sender, EventArgs e)
        {
            running = false;
            button10.Visible = false;
        }

        //Explode blocks, extra void
        private void ExplodeBlock(AcadBlockReference bRef, int blockCounter, int level)
        {
            string oldBlockName = bRef.Name;
            oldBlockName = RemoveSpecialCharacters(oldBlockName);
            try
            {
                //Get all entities in this block
                object[] explodedBlock = bRef.Explode();

                for(int i = 0; i < explodedBlock.Length; i++)
                {
                    //Set exploded block item as entity
                    AcadEntity ent = (AcadEntity)explodedBlock[i];

                    if (ent.EntityType == 7) // 7 = block
                    {
                        //If the entity is still a blockreference
                        bool layerExist = false;
                        bool frozenOrLocked = false;
                        try
                        {
                            //Set entity as AcadBlockReference
                            AcadBlockReference subBRef = ent as AcadBlockReference;
                        
                            //Explode nested blocks if checkbox is checked
                            if (checkBox9.Checked)
                            {
                                //Restart the explode void with +1 level
                                ExplodeBlock(subBRef, blockCounter, level + 1);
                                blockCounterNestedBlocks += 1;
                            }
                            else
                            {
                                //Check if the block have Solid3D entities, level1
                                bool doit = false;
                                try
                                {
                                    object[] explodedSubBlock = subBRef.Explode();
                                    foreach (AcadEntity e in explodedSubBlock)
                                    {
                                        if (e.EntityType == 3) // 3 = solid
                                        {
                                            doit = true;
                                        }
                                        if (e.EntityType == 7) // 7 = block
                                        {
                                            //Check if the block have Solid3D entities, level2
                                            AcadBlockReference subBRef2 = e as AcadBlockReference;
                                            bool doit2 = false;
                                            try
                                            {
                                                object[] explodedSubBlock2 = explodedSubBlock2 = subBRef2.Explode();
                                                foreach (AcadEntity e2 in explodedSubBlock)
                                                {
                                                    if (e2.EntityType == 3) // 3 = solid
                                                    {
                                                        doit2 = true;
                                                    }
                                                    e2.Delete();
                                                }
                                            }
                                            catch
                                            {
                                                //explode sub block is not possible for any reason... dont export this block!
                                                ignoreBlocks.Add(subBRef2.Name);
                                            }

                                            if (doit2)
                                            {
                                                doit = true;
                                            }
                                            else
                                            {
                                                subBRef2.Delete();
                                            }
                                        }
                                        e.Delete();
                                    }
                                }
                                catch
                                {
                                    //explode block is not possible for any reason... dont export this block!
                                    ignoreBlocks.Add(subBRef.Name);
                                }


                                //If not, try to change the nested block layer
                                if (doit)
                                {
                                    string oldSubBlockName = RemoveSpecialCharacters(subBRef.Name);

                                    string newLayer = separator + blockCounter.ToString() + separator + oldBlockName.ToUpper() + separator + oldSubBlockName + separator + subBRef.Layer + separator + "LEVEL-" + level;
                                    foreach (AcadLayer acLay in thisDrawing.Layers)
                                    {
                                        //Check if the layer already exist
                                        if (acLay.Name == newLayer)
                                        {
                                            layerExist = true;
                                        }

                                        //Check if the block layer is frozen or locked
                                        if (acLay.Name == subBRef.Layer)
                                        {
                                            if (acLay.Freeze || acLay.Lock || !acLay.LayerOn)
                                            {
                                                frozenOrLocked = true;
                                            }
                                        }
                                    }

                                    //If the layer not exist, nd the block layer is not frozen or locked, add it
                                    if (!layerExist && !frozenOrLocked)
                                    {
                                        thisDrawing.Layers.Add(newLayer);
                                    }
                                    //If the block layer is not frozen and not locked, apply the new layer
                                    if (!frozenOrLocked)
                                    {
                                        subBRef.Layer = newLayer;
                                    }
                                }
                                else
                                {
                                    //No solid found
                                    subBRef.Delete();
                                }
                            }
                        }
                        catch { }
                    }
                    else if (ent.EntityType == 3) // 3 = Solid3D
                    {
                        //if the entity is a Solid3d
                        if (radioButton2.Checked)
                        {
                            //Objects / block
                            string newLayer = separator + blockCounter.ToString() + separator + oldBlockName.ToUpper() + separator + ent.Layer + separator + i + separator + "LEVEL-" + level;
                            thisDrawing.Layers.Add(newLayer);
                            ent.Layer = newLayer;
                        }
                        else
                        {
                            //Layer / block
                            bool layerExist = false;
                            bool frozenOrLocked = false;
                            try
                            {
                                string newLayer = separator + blockCounter.ToString() + separator + oldBlockName.ToUpper() + separator + ent.Layer + separator + "LEVEL-" + level;
                                foreach (AcadLayer acLay in thisDrawing.Layers)
                                {
                                    //Check if the layer already exist
                                    if (acLay.Name == newLayer)
                                    {
                                        layerExist = true;
                                    }

                                    //Check if the entity layer is frozen or locked
                                    if (acLay.Name == ent.Layer)
                                    {
                                        if(acLay.Freeze || acLay.Lock || !acLay.LayerOn)
                                        {
                                            frozenOrLocked = true;
                                        }
                                    }
                                }

                                //If the layer not exist, nd the entity layer is not frozen or locked, add it
                                if (!layerExist && !frozenOrLocked)
                                {
                                    thisDrawing.Layers.Add(newLayer);
                                }
                                // If the entity layer is not frozen and not locked, apply the new layer
                                if (!frozenOrLocked)
                                {
                                    ent.Layer = newLayer;
                                }
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // entityTypes: 5=??, 19=Line, 24=?? 
                        //MessageBox.Show("Wrong entity type: " + ent.EntityType.ToString());
                    }
                }

                //If block was a nested block, delete the reference
                if (level > 0)
                {
                    bRef.Delete();
                }
            }
            catch
            {
                //explode block is not possible for any reason... dont export this block!
                ignoreBlocks.Add(bRef.Name);
                bRef.Delete();
            }
        }

        private void writeSTL(AcadGroup group, double progressStep3, string stlOutputPath)
        {
            //Remove the prefix for the filename
            string stlFilename = group.Name.Replace(Properties.Settings.Default.GROUP_PREFIX, "");
            if (stlFilename.Substring(0, 1) == separator)
            {
                //Remove the seperator
                stlFilename = stlFilename.Substring(1);
            }

            string completeFilename = Path.Combine(stlOutputPath, stlFilename + ".stl");
            textBox10.Text = stlFilename + Environment.NewLine + textBox10.Text;
            textBox10.Update();

            //Send the export command to the autocad commandline
            if (Properties.Settings.Default.STLOUT)
            {
                try
                {
                    thisDrawing.SendCommand("STLOUT" +
                        " " +
                        "g" +
                        " " +
                        group.Name +
                        Environment.NewLine +
                        " " +
                        " " +
                        completeFilename +
                        Environment.NewLine);
                }
                catch { }
            }
            else
            {
                try
                {
                    thisDrawing.SendCommand("_export" +
                        " " +
                        completeFilename +
                        Environment.NewLine +
                        "g" +
                        " " +
                        group.Name +
                        " " + " ");
                }
                catch { }
            }

            //Timeout if checkbox checked
            if (checkBox8.Checked)
            {
                int t;
                bool inum = int.TryParse(comboBox1.GetItemText(comboBox1.SelectedItem), out t);
                if (inum)
                {
                    Thread.Sleep(t);
                }
            }

            //Update progressBar
            progressBar1.Value = UpdateProgressBar(progressStep3);
            progressBar1.Refresh();
        }

        //Create a new drawing
        private void CreateNewDrawing(DocumentCollection documentManager)
        {
            try
            {
                string strTemplatePath = "acad.dwt";
                ObjectIdCollection oIdCol = new ObjectIdCollection();

                TypedValue[] tv = new TypedValue[3] {
                        new TypedValue(67, 0), //Model space
                        new TypedValue((int)DxfCode.Visibility, 0),
                        new TypedValue((int)DxfCode.Start, "INSERT,3DSOLID"),
                        };
                SelectionFilter sf = new SelectionFilter(tv);
                PromptSelectionResult res = documentManager.MdiActiveDocument.Editor.SelectAll(sf);
                SelectionSet copySset = res.Value;

                foreach (SelectedObject so in copySset)
                {
                    oIdCol.Add(so.ObjectId);
                }

                copySset = null;
                res = null;

                try
                {
                    Document acNewDoc = documentManager.Add(strTemplatePath);
                    documentManager.MdiActiveDocument = acNewDoc;
                    Database acNewDb = acNewDoc.Database;

                    using (DocumentLock acDocLock = acNewDoc.LockDocument())
                    {
                        using (Transaction tr = acNewDb.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(acNewDb.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acNewDb.CurrentSpaceId, OpenMode.ForWrite);
                            IdMapping map = new IdMapping();
                            acNewDb.WblockCloneObjects(oIdCol, btr.ObjectId, map, DuplicateRecordCloning.Ignore, false);
                            tr.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Can't copy objects to the new drawing...  Original error: " + ex.Message);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Move objects to a new drawing failed... Original error:" + ex.Message);
                return;
            }
        }

        //Get the number of vertices from a Solid3d
        private int NumberOfVertices(Entity ent)
        {
            //reference acdbmgdbrep.dll is required for Brep
            using (var brep = new Brep(ent))
            {
                return brep.Vertices.Count();
            }
        }

        //Save theXML file for C4D import
        private void SaveXML(List<AcadGroup> acGroups, List<string> blocknameForXml, List<string> filenameForXml)
        {
            int xmlElementCounter = 0;
            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", "no");

            try
            {
                XmlElement rootNode = xmlDocument.CreateElement("RootNode");
                xmlDocument.InsertBefore(xmlDeclaration, xmlDocument.DocumentElement);
                xmlDocument.AppendChild(rootNode);

                foreach(AcadGroup group in acGroups)
                {
                    //Remove the prefix from the name
                    string stlFilename = group.Name.Replace(Properties.Settings.Default.GROUP_PREFIX, "");
                    string blockNameIndexAsString = "empty";

                    //If elements in the group was a block (first character is still underline)
                    if (stlFilename.Substring(0, 1) == separator)
                    {
                        //Remove the first underline (seperator)
                        stlFilename = stlFilename.Substring(1);

                        //Get the positin for the first underline in the name
                        int charsUntilUnderline = stlFilename.IndexOf(separator);

                        //Get the block number (from start until the first underline)
                        string bIndex = stlFilename.Substring(0, charsUntilUnderline);

                        for (int i = 0; i < blocknameForXml.Count; i++)
                        {
                            if (i.ToString() == bIndex)
                            {
                                blockNameIndexAsString = blocknameForXml[i];
                            }
                        }
                    }

                    XmlElement parentNode = xmlDocument.CreateElement("Parent" + xmlElementCounter.ToString());
                    xmlDocument.DocumentElement.PrependChild(parentNode);

                    XmlElement a = xmlDocument.CreateElement("CUSTOM_PROPERTY_VALUE");
                    XmlElement b = xmlDocument.CreateElement("DRAWING_NAME");
                    XmlElement c = xmlDocument.CreateElement("BLOCKNAME");
                    XmlElement d = xmlDocument.CreateElement("STL_FILENAME");
                    XmlElement e = xmlDocument.CreateElement("SHIFT");

                    XmlText aValue = xmlDocument.CreateTextNode(customProperty);
                    XmlText bValue = xmlDocument.CreateTextNode(dwg_filename);
                    XmlText cValue = xmlDocument.CreateTextNode(blockNameIndexAsString);
                    XmlText dValue = xmlDocument.CreateTextNode(stlFilename);
                    XmlText eValue = xmlDocument.CreateTextNode(checkBox2.Checked.ToString());

                    //Append nodes to the parent node
                    parentNode.AppendChild(a);
                    parentNode.AppendChild(b);
                    parentNode.AppendChild(c);
                    parentNode.AppendChild(d);
                    parentNode.AppendChild(e);

                    a.AppendChild(aValue);
                    b.AppendChild(bValue);
                    c.AppendChild(cValue);
                    d.AppendChild(dValue);
                    e.AppendChild(eValue);


                    xmlElementCounter++;
                }
                xmlDocument.Save(Path.Combine(Properties.Settings.Default.STL_PATH, var_login + ".xml"));
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Can't save the .XML file. " + "Original error: " + ex.Message);
            }
        }

        //START EXPORT!!
        private void button1_Click(object sender, EventArgs e)
        {
            SaveSettings();

            //Clear results
            textBox10.Text = "";
            progressBar1.Value = progressBar1.Minimum;
            blockCounterNestedBlocks = 0;
            button1.Enabled = false;
            running = true;
            label12.Text = "...";
            bool catchError = false;
            DocumentCollection documentManager = acApp.DocumentManager;
            List<string> blocknameForXml = new List<string>();
            List<string> filenameForXml = new List<string>();
            ignoreBlocks.Clear();
            try
            {
                Document doc = documentManager.MdiActiveDocument;
                AcadDocument tD = (AcadDocument)doc.GetAcadDocument();
                dwg_filename = tD.Name.Replace(".dwg", "");

                //Try to get the custom property
                if (Properties.Settings.Default.USE_CUSTOM_PROPERTIES)
                {
                    try
                    {
                        tD.SummaryInfo.GetCustomByKey(Properties.Settings.Default.PROPERTY_NAME, out customProperty);
                    }
                    catch { }
                }
            }
            catch { }

            //If use a new drawing is checked
            if (Properties.Settings.Default.USE_NEW_DRAWING)
            {
                CreateNewDrawing(documentManager);
            }

            //Set variables for the active drawing
            Document document = documentManager.MdiActiveDocument;
            Editor editor = document.Editor;
            Database database = document.Database;
            thisDrawing = (AcadDocument)document.GetAcadDocument();

            //Set the new FACETRES Value, and save the current value for restore aftr export
            SaveFacetres();

            //Check STL output path
            string exportPath = textBox8.Text;
            string stlOutputPath = "";
            if (!Directory.Exists(exportPath))
            {
                MessageBox.Show("Invalid STL path. Directory don't exist. Export failed.");
                return;
            }
            else
            {
                //Check login folder for STL files. create this folder if it does not exist.
                stlOutputPath = Path.Combine(exportPath, var_login);
                if (!Directory.Exists(stlOutputPath))
                {
                    Directory.CreateDirectory(stlOutputPath);
                }
                else
                {
                    //if the folder exist, delete all existing .stl files.
                    DirectoryInfo dir = new DirectoryInfo(stlOutputPath);
                    bool showDeleteError = true;
                    foreach (FileInfo file in dir.GetFiles())
                    {
                        if (file.Extension.ToUpper() == ".STL")
                        {
                            try
                            {
                                file.Delete();
                            }
                            catch (System.Exception ex)
                            {
                                //Show Error Message only once, if a file can't be deleted
                                if(showDeleteError)
                                {
                                    MessageBox.Show("Can't delete file: " + file.Name + "  | Original error:" + ex.Message);
                                }
                                showDeleteError = false;
                            }
                        }
                    }
                }
            }

            //Set FILEDIA variable to 0, for export files with the command line
            thisDrawing.SetVariable("FILEDIA", 0);

            //Reset UCS if checkbox is checked
            if (checkBox1.Checked)
            {
                editor.CurrentUserCoordinateSystem = Matrix3d.Identity;
            }

            //Get exception layers
            string[] exceptionLayerNames = null;
            if (checkBox4.Checked)
            {
                try
                {
                    exceptionLayerNames = textBox1.Text.Split(',');
                }
                catch { }
            }

            //Move objects if checkbox is checked
            if (checkBox2.Checked)
            {
                //Points for moving
                double[] pointFrom = { 0, 0, 0 };
                double[] pointTo = { 10000, 10000, 10000 };

                //Move all selected objects
                foreach (AcadEntity ent in thisDrawing.ModelSpace)
                {
                    try
                    {
                        foreach (AcadLayer lay in thisDrawing.Layers)
                        {
                            if (lay.Name == ent.Layer && (!lay.Lock && !lay.Freeze))
                            {
                                ent.Move(pointFrom, pointTo);
                            }
                        }
                    }
                    catch
                    {
                        catchError = true;
                    }
                }
            }

            //Zomm all Objects
            Autodesk.AutoCAD.Interop.AcadApplication app = (Autodesk.AutoCAD.Interop.AcadApplication)Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
            app.ZoomExtents();

            int blocks = 0;
            //count blocks for progressbar
            try
            {
                foreach (AcadEntity ent in thisDrawing.ModelSpace)
                {
                    if (ent.EntityType == 7)
                    {
                        blocks++;
                    }
                }
            }
            catch { }

            double progressStep1 = Convert.ToDouble(progressBar1.Maximum / 3);
            if (blocks > 0)
            {
                progressStep1 = progressStep1 / blocks;
            }

            //Remove all special characters
            int layerCounter = 0;
            foreach (AcadLayer lay in thisDrawing.Layers)
            {
                try
                {
                    lay.Name = RemoveSpecialCharacters(lay.Name);
                }

                //If the Layername already exist, add the counter to the name
                catch { lay.Name = RemoveSpecialCharacters(lay.Name + "-" + layerCounter); }
            }

            List<Acad3DSolid> singleSolids = new List<Acad3DSolid>();

            //Explode blocks
            if (checkBox3.Checked)
            {
                label13.Text = "Explode blocks...";
                label13.Update();
                currentLevel = 0;
                List<AcadBlockReference> oldBrefs = new List<AcadBlockReference>();
                int blockCounter = 0;
                try
                {
                    //Get all blocks in this drawing
                    foreach (AcadEntity ent in thisDrawing.ModelSpace)
                    {
                        if (ent.EntityType == 7) // 7 = block
                        {
                            //Update progressBar
                            progressBar1.Value = UpdateProgressBar(progressStep1);
                            progressBar1.Refresh();

                            AcadBlockReference bRef = ent as AcadBlockReference;
                            foreach (IAcadLayer lay in thisDrawing.Layers)
                            {
                                //If the block layer not locked, not freezed and enabled
                                if (lay.Name == bRef.Layer && (!lay.Lock && !lay.Freeze))
                                {
                                    //Check for exception layers if box checked
                                    if (!checkBox4.Checked)
                                    {
                                        ExplodeBlock(bRef, blockCounter, currentLevel);
                                        oldBrefs.Add(bRef);
                                        blockCounter++;

                                        //Remove special characters from blockname and add the new name to the list for XML
                                        string n = bRef.Name;
                                        n = RemoveSpecialCharacters(n);
                                        blocknameForXml.Add(n);
                                    }
                                    else if (checkBox4.Checked && !exceptionLayerNames.Contains(bRef.Layer))
                                    {
                                        ExplodeBlock(bRef, blockCounter, currentLevel);
                                        oldBrefs.Add(bRef);
                                        blockCounter++;

                                        //Remove special characters from blockname and add the new name to the list for XML
                                        string n = bRef.Name;
                                        n = RemoveSpecialCharacters(n);
                                        blocknameForXml.Add(n);
                                    }
                                }
                            }
                        }
                        if (ent.EntityType == 3) // 3 = Solid3D
                        {
                            singleSolids.Add(ent as Acad3DSolid);
                        }
                     }
                }
                catch
                {
                    catchError = true;
                }

                //Total exploded blocks
                label12.Text = (blockCounter + blockCounterNestedBlocks).ToString();

                //Remove old/exploded block references
                try
                {
                    foreach (AcadBlockReference bRef in oldBrefs)
                    {
                        bRef.Delete();
                    }
                    oldBrefs.Clear();
                }
                catch { }
            }
            else
            {
                //Update progressBar
                progressBar1.Value = UpdateProgressBar(Convert.ToDouble(progressBar1.Maximum / 3));
                progressBar1.Refresh();
            }

            double progressStep2 = Convert.ToDouble(progressBar1.Maximum / 3);
            if (thisDrawing.Layers.Count > 0)
            {
                progressStep2 = progressStep2 / thisDrawing.Layers.Count;
                label13.Text = "Create groups...";
                label13.Update();
            }

            //Check all layers in this drawing, and make groups
            for (int i = 0; i < thisDrawing.Layers.Count; i++)
            {
                //Update progressBar
                progressBar1.Value =  UpdateProgressBar(progressStep2);
                progressBar1.Refresh();

                AcadLayer layer = thisDrawing.Layers.Item(i);
                if(layer.LayerOn && !layer.Lock && !layer.Freeze)
                {
                    try
                    {
                        //Group name for the current layer
                        string groupName = thisDrawing.Layers.Item(i).Name.ToUpper();
                        AcadSelectionSet sset = thisDrawing.SelectionSets.Add("SelectionToCreateGroups");

                        try
                        {
                            //Define filter for selection
                            short[] FilterType = new short[4];
                            object[] FilterValue = new object[4];

                            FilterType[0] = 67; //Model/Paper
                            FilterValue[0] = 0; //Model space = 0, Paper space = 1
                            FilterType[1] = (int)DxfCode.Visibility;
                            FilterValue[1] = 0; //Visible = 0, Invisible = 1
                            FilterType[2] = (int)DxfCode.LayerName;
                            FilterValue[2] = layer.Name;
                            FilterType[3] = (int)DxfCode.Start;
                            FilterValue[3] = "INSERT,3DSOLID";

                            object filterCode = FilterType;

                            //Select objects
                            sset.Select(AcSelect.acSelectionSetAll, null, null, filterCode, FilterValue);
                        }
                        catch { }

                        if (sset.Count > 0)
                        {
                            try
                            {
                                //Objects per Solid3D outside of blocks - checkbox checked
                                if (checkBox11.Checked)
                                {
                                    //Create group for each Solid3D on this layer
                                    List<AcadEntity> entitiesToGroup = new List<AcadEntity>();
                                    for (int entityCounter = 0; entityCounter < sset.Count; entityCounter++)
                                    {
                                        if (sset.Item(entityCounter).Layer.Substring(0, 1) != separator && sset.Item(entityCounter).EntityType == 3)
                                        {
                                            AcadEntity[] entityToGroup = new AcadEntity[1] { sset.Item(entityCounter) };
                                            AcadGroup group = thisDrawing.Groups.Add(Properties.Settings.Default.GROUP_PREFIX + sset.Item(entityCounter).Layer.ToUpper() + entityCounter);
                                            group.AppendItems(entityToGroup);
                                        }
                                        else
                                        {
                                            //if the object is not a Solid3D or was in a block, add it to the list
                                            entitiesToGroup.Add(sset.Item(entityCounter));
                                        }
                                    }

                                    //Create group for remaining objects
                                    if(entitiesToGroup.Count > 0)
                                    {
                                        try
                                        {
                                            AcadGroup group = thisDrawing.Groups.Add(Properties.Settings.Default.GROUP_PREFIX + groupName);
                                            AcadEntity[] entArray = entitiesToGroup.ToArray();
                                            group.AppendItems(entArray);
                                        }
                                        catch { }
                                    }
                                }
                                else
                                {
                                    //Add all entities in the selectionset to a array
                                    AcadEntity[] entitiesToGroup = new AcadEntity[sset.Count];
                                    for (int entityCounter = 0; entityCounter < sset.Count; entityCounter++)
                                    {
                                        entitiesToGroup[entityCounter] = sset.Item(entityCounter);
                                    }

                                    //If there are entites in the selectionSet
                                    if (entitiesToGroup.Count() > 0)
                                    {
                                        //Create group
                                        AcadGroup group = thisDrawing.Groups.Add(Properties.Settings.Default.GROUP_PREFIX + groupName);
                                        group.AppendItems(entitiesToGroup);
                                    }
                                }
                            }
                            catch
                            {
                                catchError = true;
                            }
                        }

                        //Delete the selectionset
                        try
                        {
                            sset.Delete();
                        }
                        catch { }
                    }
                    catch
                    {
                        catchError = true;
                    }
                }
            }

            List<AcadGroup> acGroups = new List<AcadGroup>();
            foreach (AcadGroup group in thisDrawing.Groups)
            {
                //Check for PREFIX string. only export groups with the correct prefix
                if (group.Name.Substring(0, Properties.Settings.Default.GROUP_PREFIX.Length) == Properties.Settings.Default.GROUP_PREFIX)
                {
                    acGroups.Add(group);
                }
            }

            //Activate the stop button
            button10.Visible = true;
            button10.Update();

            //Create files
            if (acGroups.Count > 0)
            {
                //If group not empty
                double progressStep3 = Convert.ToDouble((progressBar1.Maximum - progress) / acGroups.Count);
                label13.Text = "Write files...";
                label13.Update();

                //Write STL file (Autocad Commandline)
                foreach (AcadGroup group in acGroups)
                {
                    if (running)
                    {
                        writeSTL(group, progressStep3, stlOutputPath);
                    }
                }
            }
            string endMessage = "Export complete.";

            //Create the XML file
            if (running)
            {
                SaveXML(acGroups, blocknameForXml, filenameForXml);
            }
            else
            {
                endMessage = "Export stopped. Not complete...";
            }

            if(catchError)
            {
                endMessage = "There was one or more errors during export!\nExport maybe not complete...";
            }
            else if(ignoreBlocks.Count > 0)
            {
                endMessage = "There was a error with exploding blocks. Not all blocks are exported.\n";
                foreach(string blockName in ignoreBlocks)
                {
                    endMessage += blockName + ", ";
                }
            }

            //Restore FILEDIA variable
            thisDrawing.SetVariable("FILEDIA", 1);

            //Restore the FACETRES variable with the previous value
            RestoreFacetres();

            //Hide the Stop Button
            running = false;
            button10.Visible = false;

            //Close the new drawing if checked
            if (Properties.Settings.Default.USE_NEW_DRAWING && Properties.Settings.Default.CLOSE_NEW_DWG)
            {
                document.CloseAndDiscard();
            }

            //Info for export complete
            MessageBox.Show(endMessage);

            //Activate the export button again
            button1.Enabled = true;

            //Reset the progressBar
            progressBar1.Value = progressBar1.Minimum;
            progress = 0;
            label13.Text = "Status:";

            //Editor Command Line Value
            //List<string> history = new List<string>(Autodesk.AutoCAD.Internal.Utils.GetLastCommandLines(2, true));
            //foreach(string s in history)
            //{
            //    MessageBox.Show(s);
            //}
        }
    }
}
