using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ItemUploadTool.PD_EDWDataSetTableAdapters;
using ItemUploadTool.PD_EDWDataSet1TableAdapters;
using ExcelDataReader.Core;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Log;
using ExcelDataReader;
using System.IO;
using System.Windows.Automation;



namespace ItemUploadTool
{

    public partial class Form1 : Form
    {
        int triggerfile = 0;
        DataTable table1 = new DataTable("PartcodeString");
        DataTable table2 = new DataTable("PartcodeString");
        DataTable table3 = new DataTable("BOMLOOKUP");
        DataTable table4 = new DataTable("Refdwg");
        DataTable table5 = new DataTable("bcodeslist");
        public DataTable bom_Table;

        int trigger = 0;
        PD_EDWDataSet1TableAdapters.JDEItemMasterTableAdapter itemmaster = new PD_EDWDataSet1TableAdapters.JDEItemMasterTableAdapter();
        PD_EDWDataSet2TableAdapters.spoolsTableAdapter BOMLOOKS = new PD_EDWDataSet2TableAdapters.spoolsTableAdapter();
        PD_EDWDataSetTableAdapters.SpecTableAdapter missingcodes = new PD_EDWDataSetTableAdapters.SpecTableAdapter();
        PD_EDWDataSetTableAdapters.JDEItemMasterTableAdapter LIVEDESCCHECK = new PD_EDWDataSetTableAdapters.JDEItemMasterTableAdapter();
        BOMConnectDataSetTableAdapters.v_spoolsTableAdapter reffinder = new BOMConnectDataSetTableAdapters.v_spoolsTableAdapter();
        BOMConnectDataSetTableAdapters.MTODashboardTableAdapter bomfinder = new BOMConnectDataSetTableAdapters.MTODashboardTableAdapter();

        public Form1()
        {


            //MessageBox.Show(s);
            
            InitializeComponent();



        }


        /// <Public Variables and Functions used with the GL CLass Code Label Click and Return of data from Form2>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        public string whichtab;
        public void SetTextt1(string customText)
        {
            this.t1gl.Text = customText;
        }
        public void SetTextt2(string customText)
        {
            this.t2gl.Text = customText;
        }
        public void SetTextt1m(string customText)
        {
            this.t1mat.Text = customText;
        }
        public void SetTextt2m(string customText)
        {
            this.t2mat.Text = customText;
        }
        /// <Variables for the current Last Row Count within the DataGridViews>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        int currentlastcellt1 = 0;
        int currentlastcellt2 = 0;
        int currentlastcellt3 = 0;
        int currentlastcellt4 = 0;
        private void t1pcode_TextChanged(object sender, EventArgs e)
        {
            label1.ForeColor = Color.Black;
            label1.Text = "Status:*";
            if (t1pcode.Text == String.Empty)
            {
                label1.ForeColor = Color.Black;
                label1.Text = "Status:*";
            }

            tboxsubcom.Text = String.Empty;
            tboxcom.Text = String.Empty;
            tboxsize1.Text = String.Empty;
            tboxsize2.Text = String.Empty;
            tboxsch.Text = String.Empty;
            tboxrating.Text = String.Empty;
            tboxmat.Text = String.Empty;

            try
            {
                string pcodeval = t1pcode.Text.ToString();
                string comm = pcodeval.Substring(0, 1);
                string subcomm = "";
                string size1text = "";
                string mattext = "";
                if (comm == "O")
                {
                    subcomm = pcodeval.Substring(1, 3);
                    if (pcodeval.Length > 6)
                    {
                        size1text = pcodeval.Substring(4, 3);
                    }
                }
                else
                {
                    subcomm = pcodeval.Substring(1, 4);
                    if (pcodeval.Length > 7)
                    {
                        size1text = pcodeval.Substring(5, 3);
                    }
                }
                if (pcodeval.Length == 15) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 14) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 22) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 18) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 20) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 19) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 16) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 17) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 23) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 10) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 21) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }
                if (pcodeval.Length == 24) { mattext = pcodeval.Substring(pcodeval.Length - 3, 3); }

                tboxcom.Text = comm;
                tboxsubcom.Text = subcomm;
                tboxsize1.Text = size1text;
                tboxmat.Text = mattext;

                if (t1template.Text == "PIPE1")
                {
                    tboxsch.Text = pcodeval.Substring(8, 4);
                }

                if (t1template.Text == "BUTTWLD1")
                {
                    tboxsch.Text = pcodeval.Substring(8, 4);
                }

                if (t1template.Text == "OLET1")
                {
                    tboxsize2.Text = pcodeval.Substring(7, 3);
                    tboxsch.Text = pcodeval.Substring(10, 4);
                }

                if (t1template.Text == "BUTTWLD2")
                {
                    tboxsize2.Text = pcodeval.Substring(8, 3);
                    tboxsch.Text = pcodeval.Substring(11, 4);
                }

                if (t1template.Text == "FLANGE1")
                {
                    tboxsch.Text = pcodeval.Substring(8, 4);
                    tboxrating.Text = pcodeval.Substring(12, 4);
                }

                if (t1template.Text == "FLANGE2" || (t1template.Text == "OLET3" || (t1template.Text == "NIPPLE1" || (t1template.Text == "SWAGE1"))))
                {
                    tboxsize2.Text = pcodeval.Substring(8, 3);
                    tboxsch.Text = pcodeval.Substring(11, 4);
                    tboxrating.Text = pcodeval.Substring(15, 4);
                }

                if (t1template.Text == "FLANGE3" || (t1template.Text == "FORGING1" || (t1template.Text == "VALVE1" || (t1template.Text == "MISCMTL1" || (t1template.Text == "MISCMTL3")))))
                {
                    tboxrating.Text = pcodeval.Substring(8, 4);
                }

                if (t1template.Text == "FLANGE4" || (t1template.Text == "FORGING2"))
                {
                    tboxsize2.Text = pcodeval.Substring(8, 3);
                    tboxrating.Text = pcodeval.Substring(11, 4);

                }

                if (t1template.Text == "OLET2")
                {
                    tboxsize2.Text = pcodeval.Substring(7, 3);
                    tboxrating.Text = pcodeval.Substring(10, 4);

                }

                if (t1template.Text == "FORGING3")
                {
                    tboxsize2.Text = pcodeval.Substring(8, 3);

                }

                if (t1template.Text == "VALVE2")
                {
                    tboxmat.Text = pcodeval.Substring(pcodeval.Length - 3, 3);
                    tboxrating.Text = pcodeval.Substring(8, 4);

                }
                if (t1template.Text == String.Empty)
                {

                }





            }
            catch
            {

            }
            timer1.Stop();
            timer1.Start();

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            string s = Environment.UserName;

            s = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(s.ToLower());
            string usernameformatted = s.Replace(".", " ");
            this.Text = s;
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major.ToString() +"."+ System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build.ToString().Left(1);
            this.Text = String.Format("Item Upload Tool {0} - " + usernameformatted, version);
            // TODO: This line of code loads data into the 'pD_EDWDataSet.JDEItemMaster' table. You can move, or remove it, as needed.
            //this.jDEItemMasterTableAdapter.Fill(this.pD_EDWDataSet.JDEItemMaster);


            t4dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            // TODO: This line of code loads data into the 'pD_EDWDataSet.Spec' table. You can move, or remove it, as needed.
           // this.specTableAdapter.Fill(this.pD_EDWDataSet.Spec);
            t1template.Items.Add("PIPE1");
            t1template.Items.Add("BUTTWLD1");
            t1template.Items.Add("BUTTWLD2");
            t1template.Items.Add("FLANGE1");
            t1template.Items.Add("FLANGE2");
            t1template.Items.Add("FLANGE3");
            t1template.Items.Add("FLANGE4");
            t1template.Items.Add("FORGING1");
            t1template.Items.Add("FORGING2");
            t1template.Items.Add("FORGING3");
            t1template.Items.Add("OLET1");
            t1template.Items.Add("OLET2");
            t1template.Items.Add("OLET3");
            t1template.Items.Add("VALVE1");
            t1template.Items.Add("VALVE2");
            t1template.Items.Add("VALVE3");
            t1template.Items.Add("VALVE4");
            t1template.Items.Add("NIPPLE1");
            t1template.Items.Add("SWAGE1");
            t1template.Items.Add("MISCMTL1");
            t1template.Items.Add("MISCMTL2");
            t1template.Items.Add("MISCMTL3");
            t1template.Items.Add("MISCMTL4");
            t1template.Items.Add("MISCMTL5");
            t1template.Items.Add("MISCMTL6");
            t1template.Items.Add("MISCMTL7");
            t1template.Items.Add("MISCMTL8");



        }
        /// <Auto GL Class population first three chars for carbon and stainless>
        /// //////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t2mat_TextChanged_1(object sender, EventArgs e)
        {
            if (t2mat.Text == String.Empty)
            {
                t2gl.Text = "40";
            }
            if (t2mat.Text == "40")
            {
                t2gl.Text = "403";
            }
            if (t2mat.Text == "00")
            {
                t2gl.Text = "401";
            }
            if (t2mat.Text == "42")
            {
                t2gl.Text = "403";
            }
            if (t2mat.Text == "60")
            {
                t2gl.Text = "401";
            }
            if (t2mat.Text == "88")
            {
                t2gl.Text = "404";
            }
            if (t2mat.Text == "83")
            {
                t2gl.Text = "404";
            }
            if (t2mat.Text == "70")
            {
                t2gl.Text = "408";
            }
            t2gl.SelectionStart = 0;
            t2gl.SelectionLength = t2gl.Text.Length;



        }
        /// <Convert Description into Support Partcode Label>
        /// /////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t2convertlbl_Click(object sender, EventArgs e)
        { }
        /// <Undo Add>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1undoadd_Click(object sender, EventArgs e)
        {
            {
                for (int i = 0; i < 5; i++)
                {
                    if (dataGridView1.Rows.Count >= 2)
                    {
                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);
                    }
                }
                currentlastcellt1 = currentlastcellt1 - 5;
            }
        }
        private void t2undoadd_Click(object sender, EventArgs e)
        {
            {
                try
                {
                    for (int i = 0; i < 5; i++)
                    {
                        if (dataGridView2.Rows.Count >= 2)
                        {
                            dataGridView2.Rows.RemoveAt(dataGridView2.Rows.Count - 2);
                        }
                    }
                    currentlastcellt2 = currentlastcellt2 - 5;
                    dataGridView3.Rows.RemoveAt(dataGridView3.Rows.Count - 1);
                }
                catch
                { }
            }
        }
        /// <Undo Sort>
        /// /////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1undosort_Click(object sender, EventArgs e)
        {
            dataGridView1.Sort(dataGridView1.Columns["itemnumber"], ListSortDirection.Ascending);
        }
        private void t2undosort_Click(object sender, EventArgs e)
        {
            dataGridView2.Sort(dataGridView2.Columns["t2itemnumber"], ListSortDirection.Ascending);
        }
        /// <Copy Function>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1copy_Click(object sender, EventArgs e)
        {


            this.dataGridView1.Sort(this.dataGridView1.Columns["itemorbranch"], ListSortDirection.Ascending);
            try
            {
                for (int loop = 0; loop < currentlastcellt1; loop++)
                    dataGridView1.Rows[loop].Selected = true;
                Clipboard.SetDataObject(
                this.dataGridView1.GetClipboardContent());
            }
            catch
            {

            }
            string q = Environment.UserName;
            q = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(q.ToLower());
            string usernameformatted = q.Replace(".", " ");

            DateTime now = DateTime.Now;
            string nowdate = now.ToString();
            nowdate = nowdate.Replace(".", "").Replace("/", "").Replace(":", "").Replace(" ", "");
            string cb = Environment.UserName + System.Environment.NewLine + Clipboard.GetText();
            System.IO.File.WriteAllText("V:\\MTO\\exe tools\\Item Upload Tool\\Logs\\" + usernameformatted + nowdate + ".txt", cb);

        }
        private void t2copy_Click(object sender, EventArgs e)
        {
            this.dataGridView1.DefaultCellStyle.WrapMode =
                DataGridViewTriState.False;
            this.dataGridView2.Sort(this.dataGridView2.Columns["t2itemorbranch"], ListSortDirection.Ascending);
            try
            {
                for (int loop = 0; loop < currentlastcellt2; loop++)
                    dataGridView2.Rows[loop].Selected = true;
                Clipboard.SetDataObject(
                this.dataGridView2.GetClipboardContent());
            }
            catch { }
            string q = Environment.UserName;
            q = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(q.ToLower());
            string usernameformatted = q.Replace(".", " ");

            DateTime now = DateTime.Now;
            string nowdate = now.ToString();
            nowdate = nowdate.Replace(".", "").Replace("/", "").Replace(":", "").Replace(" ", "");
            string cb = Environment.UserName + System.Environment.NewLine + Clipboard.GetText();
            System.IO.File.WriteAllText("V:\\MTO\\exe tools\\Item Upload Tool\\Logs\\" + usernameformatted + nowdate + ".txt", cb);
        }
        /// <Clear Data>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1cleardata_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            currentlastcellt1 = 0;
        }
        private void t2cleardata_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Refresh();
            dataGridView3.Rows.Clear();
            dataGridView3.Refresh();
            currentlastcellt2 = 0;
        }
        /// <Reset>
        /// //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void t1reset_Click(object sender, EventArgs e)
        {

            t1pcode.Text = string.Empty;
            t1desc.Text = string.Empty;
            t1mat.Text = string.Empty;
            t1gl.Text = string.Empty;
            t1wt.Text = string.Empty;
            t1sa.Text = string.Empty;
            t1template.Text = String.Empty;
            label1.ForeColor = Color.Black;
            label1.Text = "Status:*";
            t1missingitems.Items.Clear();
            t1missingitems.SelectedItem = null;
        }
        private void t2reset_Click(object sender, EventArgs e)
        {
            t2pcode.Text = string.Empty;
            t2desc.Text = string.Empty;
            t2mat.Text = string.Empty;
            t2gl.Text = string.Empty;
            t2wt.Text = string.Empty;
            t2sa.Text = string.Empty;
            t2listfromclip.Items.Clear();
            t2listfromclip.SelectedItem = null;

        }
        ToolTip t = new ToolTip();
        /// <Add Data-Materials>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1add_Click(object sender, EventArgs e)
        {
            DataTable SC_table = null;
            DataTable S1_table = null;
            DataTable S2_table = null;
            DataTable SCH_table = null;
            DataTable RATING_table = null;
            DataTable SGC_table = null;
            String[] sc =null;
            String s1 = null;
            String s2 = null;
            String sch = null;
            String rtg = null;
            String sgc = null;
            if (tboxsubcom.Text != "")
            {
                SC_table = LIVEDESCCHECK.GetDataBySUBCOMM(tboxsubcom.Text);
                sc = SC_table.Rows[0]["SUBCOMM_DESC"].ToString().Split(null);
            }
                if (tboxsize1.Text != "")
                {
                    S1_table = LIVEDESCCHECK.GetDataBySIZE1(tboxsize1.Text);
                       s1 = S1_table.Rows[0]["SIZE1_DESC"].ToString();
            }
                    if (tboxsize2.Text != "")
                    {
                        S2_table = LIVEDESCCHECK.GetDataBySIZE2(tboxsize2.Text);
                         s2 = S2_table.Rows[0]["SIZE2_DESC"].ToString();
            }
                        if (tboxsch.Text != "")
                        {
                            SCH_table = LIVEDESCCHECK.GetDataBySCH(tboxsch.Text);
                             sch = SCH_table.Rows[0]["SCH_DESC"].ToString();
            }
                            if (tboxrating.Text != "")
                            {
                                RATING_table = LIVEDESCCHECK.GetDataByRATING(tboxrating.Text);
                                rtg = RATING_table.Rows[0]["RATING_EC_DESC"].ToString();
            }
                                if (tboxmat.Text != "")
                                {
                                    SGC_table = LIVEDESCCHECK.GetDataBySGC(tboxmat.Text);
                                    sgc = SGC_table.Rows[0]["SGC_DESC"].ToString();
            }

            
            
            
            
            
            
            foreach (string scom in sc)
            {
                if (!t1desc.ToString().ToLower().Contains(scom.ToLower()))

                    t.Show("Error in description - Sub-Commodity", t1desc, 5000);
                
          }

            if (!t1desc.ToString().ToLower().Contains(s1.ToLower() + " "))
            {
                t.Show("Error in description - Size 1", t1desc, 5000);
                
            }
            if (tboxsize2.Text != "")
            {
                if (!t1desc.ToString().ToLower().Contains(" " + s2.ToLower() + " "))
                {
                    t.Show("Error in description - Size 2", t1desc, 5000);
                    
                }
            }
            if (tboxsch.Text != "")
            {
                if (!t1desc.ToString().ToLower().Contains(" " + sch.ToLower() + " "))
                {
                    t.Show("Error in description - Sch", t1desc, 5000);
                    
                }
            }
            if (tboxrating.Text != "")
            {
                if (!t1desc.ToString().ToLower().Contains(" " + rtg.ToLower() + " "))
                {
                    t.Show("Error in description - Rating", t1desc, 5000);
                    
                }
            }
            if (tboxmat.Text != "")
            {
                if (!t1desc.ToString().ToLower().Contains(" " + sgc.ToLower() + " "))
                {
                    t.Show("Error in description - SGC", t1desc, 5000);
                    
                }
            }



            string LPT;
            if (t1TagItemCB.Checked)
            {
                LPT = "3";
            }
            else if (!t1TagItemCB.Checked)
            {
                LPT = "3";
            }


            if (t1template.Text == "")
            {
                MessageBox.Show("Please Select a Teamplate", "Missing Template", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string pcodecheck = "";
            string desccheck = "";
            for (int i = 1; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows.Count <= 1)
                {

                }
                else
                {
                    try
                    {
                        if (t1desc.Text.Length <= 30)
                        {
                            pcodecheck = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            desccheck = dataGridView1.Rows[i].Cells[6].Value.ToString();
                        }
                        else
                        {
                            pcodecheck = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            desccheck = dataGridView1.Rows[i].Cells[6].Value.ToString() + dataGridView1.Rows[i].Cells[7].Value.ToString();
                        }

                        if (t1pcode.Text == pcodecheck || (t1desc.Text == desccheck))
                        {
                            MessageBox.Show("This item Seems to Exist already", "Duplicate Item", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }


                    }
                    catch { }
                }
            }

            label1.ForeColor = Color.Black;
            label1.Text = "Status:*";
            try
            {
                double t1wtnum = Convert.ToDouble(t1wt.Text);
                double t1sanum = Convert.ToDouble(t1sa.Text);

                if (t1mat.Text == String.Empty || (t1gl.Text == String.Empty || (t1wt.Text == String.Empty || (t1sa.Text == String.Empty || (t1desc.Text == String.Empty || (t1pcode.Text == String.Empty || (t1wtnum <= 0 || (t1sanum <= 0) || (t1sanum >= t1wtnum))))))))
                {
                    MessageBox.Show("Please make sure all fields are filled out.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                List<string> list = new List<string>();
                list.Add("");
                list.Add("31");
                list.Add("32");
                list.Add("31500");
                list.Add("32500");
                int i = 1;
                foreach (string item in list)
                {
                    int counter = i;
                    if (i > 2)
                    {
                        counter = 2;
                    }
                    try
                    {
                        Int32 lastrow = dataGridView1.Rows.Count - 1;
                        this.dataGridView1.Rows.Add();

                        dataGridView1.Rows[lastrow].Cells[0].Value = "02";
                        dataGridView1.Rows[lastrow].Cells[1].Value = counter;
                        dataGridView1.Rows[lastrow].Cells[2].Value = item;

                        dataGridView1.Rows[lastrow].Cells[4].Value = t1pcode.Text;
                        dataGridView1.Rows[lastrow].Cells[5].Value = t1pcode.Text;


                        string alldesc = t1desc.Text.ToString();


                        dataGridView1.Rows[lastrow].Cells[9].Value = tboxcom.Text;
                        dataGridView1.Rows[lastrow].Cells[10].Value = tboxsubcom.Text;
                        dataGridView1.Rows[lastrow].Cells[11].Value = tboxsize1.Text;
                        dataGridView1.Rows[lastrow].Cells[12].Value = tboxsize2.Text;
                        dataGridView1.Rows[lastrow].Cells[13].Value = tboxsch.Text;
                        dataGridView1.Rows[lastrow].Cells[14].Value = tboxrating.Text;
                        dataGridView1.Rows[lastrow].Cells[15].Value = tboxmat.Text;
                        dataGridView1.Rows[lastrow].Cells[32].Value = t1template.Text;
                        if (t1desc.Text.Length >= 30)
                        {
                            dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text.Substring(0, 30);
                            dataGridView1.Rows[lastrow].Cells[7].Value = t1desc.Text.Substring(30, (t1desc.Text.Length - 30));
                        }

                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text.Substring(0, t1desc.Text.Length);
                        }

                        dataGridView1.Rows[lastrow].Cells[17].Value = t1mat.Text;
                        dataGridView1.Rows[lastrow].Cells[24].Value = t1gl.Text;
                        dataGridView1.Rows[lastrow].Cells[30].Value = t1wt.Text;
                        dataGridView1.Rows[lastrow].Cells[31].Value = t1sa.Text;

                        if (tboxcom.Text == "P")
                        {
                            dataGridView1.Rows[lastrow].Cells[21].Value = "FT";
                        }
                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[21].Value = "EA";
                        }
                        dataGridView1.Rows[lastrow].Cells[22].Value = "LB";
                        dataGridView1.Rows[lastrow].Cells[23].Value = "SF";
                        if (tboxcom.Text == "V" || tboxcom.Text == "M" || t1TagItemCB.Checked)
                        {
                            dataGridView1.Rows[lastrow].Cells[25].Value = "3";
                        }
                        else if (tboxcom.Text != "V" || tboxcom.Text != "M" || !t1TagItemCB.Checked)
                        {
                            dataGridView1.Rows[lastrow].Cells[25].Value = "2";
                        }
                        dataGridView1.Rows[lastrow].Cells[26].Value = "0";
                        dataGridView1.Rows[lastrow].Cells[27].Value = "P";
                        dataGridView1.Rows[lastrow].Cells[28].Value = "S";
                        dataGridView1.Rows[lastrow].Cells[29].Value = "3650";
                        dataGridView1.Rows[lastrow].Cells[33].Value = "Y";
                        i++;
                    }

                    catch { }
                }
                currentlastcellt1 = currentlastcellt1 + 5;
            }
            catch
            { }
        }

        private void addAsCorrection04ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            string pcodecheck = "";
            string desccheck = "";
            for (int i = 1; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows.Count <= 1)
                {

                }
                else
                {
                    try
                    {
                        if (t1desc.Text.Length <= 30)
                        {
                            pcodecheck = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            desccheck = dataGridView1.Rows[i].Cells[6].Value.ToString();
                        }
                        else
                        {
                            pcodecheck = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            desccheck = dataGridView1.Rows[i].Cells[6].Value.ToString() + dataGridView1.Rows[i].Cells[7].Value.ToString();
                        }

                        if (t1pcode.Text == pcodecheck || (t1desc.Text == desccheck))
                        {
                            MessageBox.Show("This item Seems to Exist already", "Duplicate Item", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    catch { }
                }
            }
            label1.ForeColor = Color.Black;
            label1.Text = "Status:*";
            try
            {
                double t1wtnum = Convert.ToDouble(t1wt.Text);
                double t1sanum = Convert.ToDouble(t1sa.Text);

                if (t1mat.Text == String.Empty || (t1gl.Text == String.Empty || (t1wt.Text == String.Empty || (t1sa.Text == String.Empty || (t1desc.Text == String.Empty || (t1pcode.Text == String.Empty || (t1wtnum <= 0 || (t1sanum <= 0) || (t1sanum >= t1wtnum))))))))
                {
                    MessageBox.Show("Please make sure all fields are filled out.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                List<string> list = new List<string>();
                list.Add("");
                list.Add("31");
                list.Add("32");
                list.Add("31500");
                list.Add("32500");
                int i = 1;
                foreach (string item in list)
                {
                    int counter = i;
                    if (i > 2)
                    {
                        counter = 2;
                    }
                    try
                    {
                        Int32 lastrow = dataGridView1.Rows.Count - 1;
                        this.dataGridView1.Rows.Add();

                        dataGridView1.Rows[lastrow].Cells[0].Value = "04";
                        dataGridView1.Rows[lastrow].Cells[1].Value = counter;
                        dataGridView1.Rows[lastrow].Cells[2].Value = item;

                        dataGridView1.Rows[lastrow].Cells[4].Value = t1pcode.Text;
                        dataGridView1.Rows[lastrow].Cells[5].Value = t1pcode.Text;


                        string alldesc = t1desc.Text.ToString();


                        dataGridView1.Rows[lastrow].Cells[9].Value = tboxcom.Text;
                        dataGridView1.Rows[lastrow].Cells[10].Value = tboxsubcom.Text;
                        dataGridView1.Rows[lastrow].Cells[11].Value = tboxsize1.Text;
                        dataGridView1.Rows[lastrow].Cells[12].Value = tboxsize2.Text;
                        dataGridView1.Rows[lastrow].Cells[13].Value = tboxsch.Text;
                        dataGridView1.Rows[lastrow].Cells[14].Value = tboxrating.Text;
                        dataGridView1.Rows[lastrow].Cells[15].Value = tboxmat.Text;
                        dataGridView1.Rows[lastrow].Cells[32].Value = t1template.Text;
                        if (t1desc.Text.Length >= 30)
                        {
                            dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text.Substring(0, 30);
                            dataGridView1.Rows[lastrow].Cells[7].Value = t1desc.Text.Substring(30, (t1desc.Text.Length - 30));
                        }

                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text.Substring(0, t1desc.Text.Length);
                        }

                        dataGridView1.Rows[lastrow].Cells[17].Value = t1mat.Text;
                        dataGridView1.Rows[lastrow].Cells[24].Value = t1gl.Text;
                        dataGridView1.Rows[lastrow].Cells[30].Value = t1wt.Text;
                        dataGridView1.Rows[lastrow].Cells[31].Value = t1sa.Text;

                        if (tboxcom.Text == "P")
                        {
                            dataGridView1.Rows[lastrow].Cells[21].Value = "FT";
                        }
                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[21].Value = "EA";
                        }
                        dataGridView1.Rows[lastrow].Cells[22].Value = "LB";
                        dataGridView1.Rows[lastrow].Cells[23].Value = "SF";
                        if (tboxcom.Text == "V" || t1TagItemCB.Checked)
                        {
                            dataGridView1.Rows[lastrow].Cells[25].Value = "3";
                        }
                        if (tboxcom.Text != "V" || !t1TagItemCB.Checked)
                        {
                            dataGridView1.Rows[lastrow].Cells[25].Value = "2";
                        }
                        dataGridView1.Rows[lastrow].Cells[26].Value = "0";
                        dataGridView1.Rows[lastrow].Cells[27].Value = "P";
                        dataGridView1.Rows[lastrow].Cells[28].Value = "S";
                        dataGridView1.Rows[lastrow].Cells[29].Value = "3650";
                        dataGridView1.Rows[lastrow].Cells[33].Value = "Y";
                        i++;
                    }
                    catch { }
                }
                currentlastcellt1 = currentlastcellt1 + 5;
            }
            catch
            { }
        }

        /// <Add Data-Supports>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t2add_Click(object sender, EventArgs e)
        {
            try
            {
                double t2wtnum = Convert.ToDouble(t2wt.Text);
                double t2sanum = Convert.ToDouble(t2sa.Text);

                if (t2mat.Text == String.Empty || (t2mat.Text == "NA" || (t2gl.Text == String.Empty || (t2wt.Text == String.Empty || (t2sa.Text == String.Empty || (t2desc.Text == String.Empty || (t2pcode.Text == String.Empty || (t2wtnum <= 0 || (t2sanum <= 0) || (t2sanum >= t2wtnum)))))))))
                {
                    MessageBox.Show("Please make sure all fields are filled out.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                string pcodecheck = "";
                string desccheck = "";
                for (int K = 1; K < dataGridView2.Rows.Count; K++)
                {
                    if (dataGridView2.Rows.Count <= 1)
                    {

                    }
                    else
                    {
                        try
                        {
                            if (t2desc.Text.Length <= 30)
                            {
                                pcodecheck = dataGridView2.Rows[K].Cells[4].Value.ToString();
                                desccheck = dataGridView2.Rows[K].Cells[6].Value.ToString();
                            }
                            else
                            {
                                pcodecheck = dataGridView2.Rows[K].Cells[4].Value.ToString();
                                desccheck = dataGridView2.Rows[K].Cells[6].Value.ToString() + dataGridView2.Rows[K].Cells[7].Value.ToString();
                            }

                            if (t2pcode.Text == pcodecheck || (t2desc.Text == desccheck))
                            {
                                MessageBox.Show("This item Seems to Exist already", "Duplicate Item", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }


                        }
                        catch { }
                    }
                }


                List<string> list = new List<string>();
                list.Add("");
                list.Add("31");
                list.Add("32");
                list.Add("31500");
                list.Add("32500");
                int i = 1;
                foreach (string item in list)
                {
                    int counter = i;
                    if (i > 2)
                    {
                        counter = 2;
                    }
                    Int32 lastrow = dataGridView2.Rows.Count - 1;
                    this.dataGridView2.Rows.Add();

                    dataGridView2.Rows[lastrow].Cells[0].Value = "02";
                    dataGridView2.Rows[lastrow].Cells[1].Value = counter;
                    dataGridView2.Rows[lastrow].Cells[2].Value = item;

                    dataGridView2.Rows[lastrow].Cells[4].Value = t2pcode.Text;
                    dataGridView2.Rows[lastrow].Cells[5].Value = t2pcode.Text;
                    dataGridView2.Rows[lastrow].Cells[6].Value = t2desc.Text;
                    dataGridView2.Rows[lastrow].Cells[17].Value = t2mat.Text;
                    dataGridView2.Rows[lastrow].Cells[24].Value = t2gl.Text;
                    dataGridView2.Rows[lastrow].Cells[30].Value = t2wt.Text;
                    dataGridView2.Rows[lastrow].Cells[31].Value = t2sa.Text;

                    dataGridView2.Rows[lastrow].Cells[9].Value = "A";
                    dataGridView2.Rows[lastrow].Cells[21].Value = "EA";
                    dataGridView2.Rows[lastrow].Cells[22].Value = "LB";
                    dataGridView2.Rows[lastrow].Cells[23].Value = "SF";
                    dataGridView2.Rows[lastrow].Cells[25].Value = "2";
                    dataGridView2.Rows[lastrow].Cells[26].Value = "0";
                    dataGridView2.Rows[lastrow].Cells[27].Value = "P";
                    dataGridView2.Rows[lastrow].Cells[28].Value = "S";
                    dataGridView2.Rows[lastrow].Cells[29].Value = "3650";
                    dataGridView2.Rows[lastrow].Cells[33].Value = "Y";
                    i++;

                }
                Int32 lastrow2 = dataGridView3.Rows.Count;
                if (t2reqcheckbox.Checked)
                {
                    this.dataGridView3.Rows.Add();
                    dataGridView3.Rows[lastrow2].Cells[0].Value = t2pcode.Text;
                    dataGridView3.Rows[lastrow2].Cells[1].Value = t2qtytextbox.Text;
                    dataGridView3.Rows[lastrow2].Cells[2].Value = t2jobnumbertextbox.Text;
                    dataGridView3.Rows[lastrow2].Cells[3].Value = t2reftextbox.Text;

                }
                currentlastcellt2 = currentlastcellt2 + 5;
            }
            catch
            { }

        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                double t2wtnum = Convert.ToDouble(t2wt.Text);
                double t2sanum = Convert.ToDouble(t2sa.Text);

                if (t2mat.Text == String.Empty || (t2gl.Text == String.Empty || (t2wt.Text == String.Empty || (t2sa.Text == String.Empty || (t2desc.Text == String.Empty || (t2pcode.Text == String.Empty || (t2wtnum <= 0 || (t2sanum <= 0) || (t2sanum >= t2wtnum))))))))
                {
                    MessageBox.Show("Please make sure all fields are filled out.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                List<string> list = new List<string>();
                list.Add("");
                list.Add("31");
                list.Add("32");
                list.Add("31500");
                list.Add("32500");
                int i = 1;
                foreach (string item in list)
                {
                    int counter = i;
                    if (i > 2)
                    {
                        counter = 2;
                    }
                    Int32 lastrow = dataGridView2.Rows.Count - 1;
                    this.dataGridView2.Rows.Add();

                    dataGridView2.Rows[lastrow].Cells[0].Value = "04";
                    dataGridView2.Rows[lastrow].Cells[1].Value = counter;
                    dataGridView2.Rows[lastrow].Cells[2].Value = item;

                    dataGridView2.Rows[lastrow].Cells[4].Value = t2pcode.Text;
                    dataGridView2.Rows[lastrow].Cells[5].Value = t2pcode.Text;
                    dataGridView2.Rows[lastrow].Cells[6].Value = t2desc.Text;
                    dataGridView2.Rows[lastrow].Cells[17].Value = t2mat.Text;
                    dataGridView2.Rows[lastrow].Cells[24].Value = t2gl.Text;
                    dataGridView2.Rows[lastrow].Cells[30].Value = t2wt.Text;
                    dataGridView2.Rows[lastrow].Cells[31].Value = t2sa.Text;

                    dataGridView2.Rows[lastrow].Cells[9].Value = "A";
                    dataGridView2.Rows[lastrow].Cells[21].Value = "EA";
                    dataGridView2.Rows[lastrow].Cells[22].Value = "LB";
                    dataGridView2.Rows[lastrow].Cells[23].Value = "SF";
                    dataGridView2.Rows[lastrow].Cells[25].Value = "2";
                    dataGridView2.Rows[lastrow].Cells[26].Value = "0";
                    dataGridView2.Rows[lastrow].Cells[27].Value = "P";
                    dataGridView2.Rows[lastrow].Cells[28].Value = "S";
                    dataGridView2.Rows[lastrow].Cells[29].Value = "3650";
                    dataGridView2.Rows[lastrow].Cells[33].Value = "Y";
                    i++;
                }
                currentlastcellt2 = currentlastcellt2 + 5;
            }
            catch
            { }
        }


        /// <tab 1 Add Clear>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1addclear_Click(object sender, EventArgs e)
        {
            string pcodecheck = "";
            string desccheck = "";
            for (int i = 1; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows.Count <= 1)
                {

                }
                else
                {
                    try
                    {
                        if (t1desc.Text.Length <= 30)
                        {
                            pcodecheck = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            desccheck = dataGridView1.Rows[i].Cells[6].Value.ToString();
                        }
                        else
                        {
                            pcodecheck = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            desccheck = dataGridView1.Rows[i].Cells[6].Value.ToString() + dataGridView1.Rows[i].Cells[7].Value.ToString();
                        }

                        if (t1pcode.Text == pcodecheck || (t1desc.Text == desccheck))
                        {
                            MessageBox.Show("This item Seems to Exist already", "Duplicate Item", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    catch { }
                }
            }

            label1.ForeColor = Color.Black;
            label1.Text = "Status:*";
            try
            {
                double t1wtnum = Convert.ToDouble(t1wt.Text);
                double t1sanum = Convert.ToDouble(t1sa.Text);

                if (t1mat.Text == String.Empty || (t1gl.Text == String.Empty || (t1wt.Text == String.Empty || (t1sa.Text == String.Empty || (t1desc.Text == String.Empty || (t1pcode.Text == String.Empty || (t1wtnum <= 0 || (t1sanum <= 0) || (t1sanum >= t1wtnum))))))))
                {
                    MessageBox.Show("Please make sure all fields are filled out.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                List<string> list = new List<string>();
                list.Add("");
                list.Add("31");
                list.Add("32");
                list.Add("31500");
                list.Add("32500");
                int i = 1;
                foreach (string item in list)
                {
                    int counter = i;
                    if (i > 2)
                    {
                        counter = 2;
                    }
                    try
                    {
                        Int32 lastrow = dataGridView1.Rows.Count - 1;
                        this.dataGridView1.Rows.Add();

                        dataGridView1.Rows[lastrow].Cells[0].Value = "02";
                        dataGridView1.Rows[lastrow].Cells[1].Value = counter;
                        dataGridView1.Rows[lastrow].Cells[2].Value = item;

                        dataGridView1.Rows[lastrow].Cells[4].Value = t1pcode.Text;
                        dataGridView1.Rows[lastrow].Cells[5].Value = t1pcode.Text;


                        string alldesc = t1desc.Text.ToString();


                        dataGridView1.Rows[lastrow].Cells[9].Value = tboxcom.Text;
                        dataGridView1.Rows[lastrow].Cells[10].Value = tboxsubcom.Text;
                        dataGridView1.Rows[lastrow].Cells[11].Value = tboxsize1.Text;
                        dataGridView1.Rows[lastrow].Cells[12].Value = tboxsize2.Text;
                        dataGridView1.Rows[lastrow].Cells[13].Value = tboxsch.Text;
                        dataGridView1.Rows[lastrow].Cells[14].Value = tboxrating.Text;
                        dataGridView1.Rows[lastrow].Cells[15].Value = tboxmat.Text;
                        dataGridView1.Rows[lastrow].Cells[32].Value = t1template.Text;
                        if (t1desc.Text.Length >= 30)
                        {
                            dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text.Substring(0, 30);
                            dataGridView1.Rows[lastrow].Cells[7].Value = t1desc.Text.Substring(30, (t1desc.Text.Length - 30));
                        }

                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text.Substring(0, t1desc.Text.Length);
                        }

                        dataGridView1.Rows[lastrow].Cells[17].Value = t1mat.Text;
                        dataGridView1.Rows[lastrow].Cells[24].Value = t1gl.Text;
                        dataGridView1.Rows[lastrow].Cells[30].Value = t1wt.Text;
                        dataGridView1.Rows[lastrow].Cells[31].Value = t1sa.Text;

                        if (tboxcom.Text == "P")
                        {
                            dataGridView1.Rows[lastrow].Cells[21].Value = "FT";
                        }
                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[21].Value = "EA";
                        }
                        dataGridView1.Rows[lastrow].Cells[22].Value = "LB";
                        dataGridView1.Rows[lastrow].Cells[23].Value = "SF";
                        if (tboxcom.Text == "V")
                        {
                            dataGridView1.Rows[lastrow].Cells[25].Value = "3";
                        }
                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[25].Value = "2";
                        }
                        dataGridView1.Rows[lastrow].Cells[26].Value = "0";
                        dataGridView1.Rows[lastrow].Cells[27].Value = "P";
                        dataGridView1.Rows[lastrow].Cells[28].Value = "S";
                        dataGridView1.Rows[lastrow].Cells[29].Value = "3650";
                        dataGridView1.Rows[lastrow].Cells[33].Value = "Y";
                        i++;
                    }
                    catch { }
                }
                currentlastcellt1 = currentlastcellt1 + 5;
                t1pcode.Text = string.Empty;
                t1desc.Text = string.Empty;
                t1mat.Text = string.Empty;
                t1gl.Text = string.Empty;
                t1wt.Text = string.Empty;
                t1sa.Text = string.Empty;
            }
            catch
            { }
        }

        private void addAsCorrection04ToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            string pcodecheck = "";
            string desccheck = "";
            for (int i = 1; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows.Count <= 1)
                {

                }
                else
                {
                    try
                    {
                        if (t1desc.Text.Length <= 30)
                        {
                            pcodecheck = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            desccheck = dataGridView1.Rows[i].Cells[6].Value.ToString();
                        }
                        else
                        {
                            pcodecheck = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            desccheck = dataGridView1.Rows[i].Cells[6].Value.ToString() + dataGridView1.Rows[i].Cells[7].Value.ToString();
                        }

                        if (t1pcode.Text == pcodecheck || (t1desc.Text == desccheck))
                        {
                            MessageBox.Show("This item Seems to Exist already", "Duplicate Item", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    catch { }
                }
            }

            label1.ForeColor = Color.Black;
            label1.Text = "Status:*";
            try
            {
                double t1wtnum = Convert.ToDouble(t1wt.Text);
                double t1sanum = Convert.ToDouble(t1sa.Text);

                if (t1mat.Text == String.Empty || (t1gl.Text == String.Empty || (t1wt.Text == String.Empty || (t1sa.Text == String.Empty || (t1desc.Text == String.Empty || (t1pcode.Text == String.Empty || (t1wtnum <= 0 || (t1sanum <= 0) || (t1sanum >= t1wtnum))))))))
                {
                    MessageBox.Show("Please make sure all fields are filled out.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                List<string> list = new List<string>();
                list.Add("");
                list.Add("31");
                list.Add("32");
                list.Add("31500");
                list.Add("32500");
                int i = 1;
                foreach (string item in list)
                {
                    int counter = i;
                    if (i > 2)
                    {
                        counter = 2;
                    }
                    try
                    {
                        Int32 lastrow = dataGridView1.Rows.Count - 1;
                        this.dataGridView1.Rows.Add();

                        dataGridView1.Rows[lastrow].Cells[0].Value = "04";
                        dataGridView1.Rows[lastrow].Cells[1].Value = counter;
                        dataGridView1.Rows[lastrow].Cells[2].Value = item;

                        dataGridView1.Rows[lastrow].Cells[4].Value = t1pcode.Text;
                        dataGridView1.Rows[lastrow].Cells[5].Value = t1pcode.Text;


                        string alldesc = t1desc.Text.ToString();


                        dataGridView1.Rows[lastrow].Cells[9].Value = tboxcom.Text;
                        dataGridView1.Rows[lastrow].Cells[10].Value = tboxsubcom.Text;
                        dataGridView1.Rows[lastrow].Cells[11].Value = tboxsize1.Text;
                        dataGridView1.Rows[lastrow].Cells[12].Value = tboxsize2.Text;
                        dataGridView1.Rows[lastrow].Cells[13].Value = tboxsch.Text;
                        dataGridView1.Rows[lastrow].Cells[14].Value = tboxrating.Text;
                        dataGridView1.Rows[lastrow].Cells[15].Value = tboxmat.Text;
                        dataGridView1.Rows[lastrow].Cells[32].Value = t1template.Text;
                        if (t1desc.Text.Length >= 30)
                        {
                            dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text.Substring(0, 30);
                            dataGridView1.Rows[lastrow].Cells[7].Value = t1desc.Text.Substring(30, (t1desc.Text.Length - 30));
                        }

                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text.Substring(0, t1desc.Text.Length);
                        }

                        dataGridView1.Rows[lastrow].Cells[17].Value = t1mat.Text;
                        dataGridView1.Rows[lastrow].Cells[24].Value = t1gl.Text;
                        dataGridView1.Rows[lastrow].Cells[30].Value = t1wt.Text;
                        dataGridView1.Rows[lastrow].Cells[31].Value = t1sa.Text;

                        if (tboxcom.Text == "P")
                        {
                            dataGridView1.Rows[lastrow].Cells[21].Value = "FT";
                        }
                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[21].Value = "EA";
                        }
                        dataGridView1.Rows[lastrow].Cells[22].Value = "LB";
                        dataGridView1.Rows[lastrow].Cells[23].Value = "SF";
                        if (tboxcom.Text == "V")
                        {
                            dataGridView1.Rows[lastrow].Cells[25].Value = "3";
                        }
                        else
                        {
                            dataGridView1.Rows[lastrow].Cells[25].Value = "2";
                        }
                        dataGridView1.Rows[lastrow].Cells[26].Value = "0";
                        dataGridView1.Rows[lastrow].Cells[27].Value = "P";
                        dataGridView1.Rows[lastrow].Cells[28].Value = "S";
                        dataGridView1.Rows[lastrow].Cells[29].Value = "3650";
                        dataGridView1.Rows[lastrow].Cells[33].Value = "Y";
                        i++;
                    }
                    catch { }
                }
                currentlastcellt1 = currentlastcellt1 + 5;
                t1pcode.Text = string.Empty;
                t1desc.Text = string.Empty;
                t1mat.Text = string.Empty;
                t1gl.Text = string.Empty;
                t1wt.Text = string.Empty;
                t1sa.Text = string.Empty;
            }
            catch
            { }
        }


        /// <tab 2 Add Clear>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t2addclear_Click(object sender, EventArgs e)
        {
            try
            {
                double t2wtnum = Convert.ToDouble(t2wt.Text);
                double t2sanum = Convert.ToDouble(t2sa.Text);

                if (t2mat.Text == String.Empty || (t2gl.Text == String.Empty || (t2wt.Text == String.Empty || (t2sa.Text == String.Empty || (t2desc.Text == String.Empty || (t2pcode.Text == String.Empty || (t2wtnum <= 0 || (t2sanum <= 0) || (t2sanum >= t2wtnum))))))))
                {
                    MessageBox.Show("Please make sure all fields are filled out.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                List<string> list = new List<string>();
                list.Add("");
                list.Add("31");
                list.Add("32");
                list.Add("31500");
                list.Add("32500");
                int i = 1;
                foreach (string item in list)
                {
                    int counter = i;
                    if (i > 2)
                    {
                        counter = 2;
                    }
                    Int32 lastrow = dataGridView2.Rows.Count - 1;
                    this.dataGridView2.Rows.Add();

                    dataGridView2.Rows[lastrow].Cells[0].Value = "02";
                    dataGridView2.Rows[lastrow].Cells[1].Value = counter;
                    dataGridView2.Rows[lastrow].Cells[2].Value = item;

                    dataGridView2.Rows[lastrow].Cells[4].Value = t2pcode.Text;
                    dataGridView2.Rows[lastrow].Cells[5].Value = t2pcode.Text;
                    dataGridView2.Rows[lastrow].Cells[6].Value = t2desc.Text;
                    dataGridView2.Rows[lastrow].Cells[17].Value = t2mat.Text;
                    dataGridView2.Rows[lastrow].Cells[24].Value = t2gl.Text;
                    dataGridView2.Rows[lastrow].Cells[30].Value = t2wt.Text;
                    dataGridView2.Rows[lastrow].Cells[31].Value = t2sa.Text;

                    dataGridView2.Rows[lastrow].Cells[9].Value = "A";
                    dataGridView2.Rows[lastrow].Cells[21].Value = "EA";
                    dataGridView2.Rows[lastrow].Cells[22].Value = "LB";
                    dataGridView2.Rows[lastrow].Cells[23].Value = "SF";
                    dataGridView2.Rows[lastrow].Cells[25].Value = "2";
                    dataGridView2.Rows[lastrow].Cells[26].Value = "0";
                    dataGridView2.Rows[lastrow].Cells[27].Value = "P";
                    dataGridView2.Rows[lastrow].Cells[28].Value = "S";
                    dataGridView2.Rows[lastrow].Cells[29].Value = "3650";
                    dataGridView2.Rows[lastrow].Cells[33].Value = "Y";
                    i++;
                }
                currentlastcellt2 = currentlastcellt2 + 5;
                t2pcode.Text = string.Empty;
                t2desc.Text = string.Empty;
                t2mat.Text = string.Empty;
                t2gl.Text = string.Empty;
                t2wt.Text = string.Empty;
                t2sa.Text = string.Empty;
            }
            catch
            { }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                double t2wtnum = Convert.ToDouble(t2wt.Text);
                double t2sanum = Convert.ToDouble(t2sa.Text);

                if (t2mat.Text == String.Empty || (t2gl.Text == String.Empty || (t2wt.Text == String.Empty || (t2sa.Text == String.Empty || (t2desc.Text == String.Empty || (t2pcode.Text == String.Empty || (t2wtnum <= 0 || (t2sanum <= 0) || (t2sanum >= t2wtnum))))))))
                {
                    MessageBox.Show("Please make sure all fields are filled out.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                List<string> list = new List<string>();
                list.Add("");
                list.Add("31");
                list.Add("32");
                list.Add("31500");
                list.Add("32500");
                int i = 1;
                foreach (string item in list)
                {
                    int counter = i;
                    if (i > 2)
                    {
                        counter = 2;
                    }
                    Int32 lastrow = dataGridView2.Rows.Count - 1;
                    this.dataGridView2.Rows.Add();

                    dataGridView2.Rows[lastrow].Cells[0].Value = "04";
                    dataGridView2.Rows[lastrow].Cells[1].Value = counter;
                    dataGridView2.Rows[lastrow].Cells[2].Value = item;

                    dataGridView2.Rows[lastrow].Cells[4].Value = t2pcode.Text;
                    dataGridView2.Rows[lastrow].Cells[5].Value = t2pcode.Text;
                    dataGridView2.Rows[lastrow].Cells[6].Value = t2desc.Text;
                    dataGridView2.Rows[lastrow].Cells[17].Value = t2mat.Text;
                    dataGridView2.Rows[lastrow].Cells[24].Value = t2gl.Text;
                    dataGridView2.Rows[lastrow].Cells[30].Value = t2wt.Text;
                    dataGridView2.Rows[lastrow].Cells[31].Value = t2sa.Text;

                    dataGridView2.Rows[lastrow].Cells[9].Value = "A";
                    dataGridView2.Rows[lastrow].Cells[21].Value = "EA";
                    dataGridView2.Rows[lastrow].Cells[22].Value = "LB";
                    dataGridView2.Rows[lastrow].Cells[23].Value = "SF";
                    dataGridView2.Rows[lastrow].Cells[25].Value = "2";
                    dataGridView2.Rows[lastrow].Cells[26].Value = "0";
                    dataGridView2.Rows[lastrow].Cells[27].Value = "P";
                    dataGridView2.Rows[lastrow].Cells[28].Value = "S";
                    dataGridView2.Rows[lastrow].Cells[29].Value = "3650";
                    dataGridView2.Rows[lastrow].Cells[33].Value = "Y";
                    i++;
                }
                currentlastcellt2 = currentlastcellt2 + 5;
                t2pcode.Text = string.Empty;
                t2desc.Text = string.Empty;
                t2mat.Text = string.Empty;
                t2gl.Text = string.Empty;
                t2wt.Text = string.Empty;
                t2sa.Text = string.Empty;
            }
            catch
            { }
        }




        /// <Unhighlights GL textbox string and sets cursor position>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t2gl_GotFocus(object sender, EventArgs e)
        {
            if (t2gl.Text.Length == 0)
            {
                t2gl.SelectionStart = t2gl.Text.Length;
            }
            if (t2gl.Text.Length == 4)
            {
                t2gl.SelectionStart = 0;
                t2gl.SelectionLength = t2gl.Text.Length;
            }
            else
            {
                t2gl.SelectionStart = t2gl.Text.Length;
            }
        }
        /// <GL Class Code Label Click tab2>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t2gllbl_Click(object sender, EventArgs e)
        {
            whichtab = ("t2");
            Form2 mySecondForm = new Form2();
            mySecondForm.Opener = this;
            mySecondForm.Show();
        }
        /// <GL Class Code Label Click tab 1>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1gllbl_Click(object sender, EventArgs e)
        {
            whichtab = ("t1");
            Form2 mySecondForm = new Form2();
            mySecondForm.Opener = this;
            mySecondForm.Show();
        }
        /// <tab 2 convert desc to partcode>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string desctospcode = t2desc.Text;
            desctospcode = desctospcode.Replace(" ", "").Replace("-", "").Replace(".", "").Replace("*", "").Replace("/", "").Replace("\\", "").Replace("#", "").Replace("(", "").Replace(")", "").Replace("'", "").Replace("\"", "").Replace(";", "").Replace(":", "");
            t2pcode.Text = ("S" + desctospcode);
        }
        /// <MAT Code Label Click tab2>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>        
        private void t2matlbl_Click(object sender, EventArgs e)
        {
            whichtab = ("t2");
            Form3 mySecondForm = new Form3();
            mySecondForm.Opener = this;
            mySecondForm.Show();
        }
        /// <MAT Code Label Click tab1>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1matlbl_Click(object sender, EventArgs e)
        {
            whichtab = ("t1");
            Form3 mySecondForm = new Form3();
            mySecondForm.Opener = this;
            mySecondForm.Show();
        }
        /// <template selection tab 1>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (tboxsubcom.Text == "SMLS") { t1template.Text = ("PIPE1"); }
            if (tboxsubcom.Text == "ERWD") { t1template.Text = ("PIPE1"); }
            if (tboxsubcom.Text == "ERWX") { t1template.Text = ("PIPE1"); }
            if (tboxsubcom.Text == "EFWD") { t1template.Text = ("PIPE1"); }
            if (tboxsubcom.Text == "EFWX") { t1template.Text = ("PIPE1"); }
            if (tboxsubcom.Text == "DSAW") { t1template.Text = ("PIPE1"); }
            if (tboxsubcom.Text == "SPRW") { t1template.Text = ("PIPE1"); }
            if (tboxsubcom.Text == "SAWW") { t1template.Text = ("PIPE1"); }
            if (tboxsubcom.Text == "S45L") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "S45S") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "S90L") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "S90S") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "SCAP") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "SELH") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "SSEA") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "SSEB") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "SSEL") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "STEE") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "SCTE") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "SCRO") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "SLAT") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "S18R") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "S18S") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "W45L") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "W45S") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "W90L") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "W90S") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WCAP") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WELH") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WSEA") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WSEB") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WSEL") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WTEE") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WCTE") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WCRO") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "WLAT") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "W18R") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "X45L") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "X45S") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "X90L") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "X90S") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XCAP") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XELH") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XSEA") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XSEB") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XSEL") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XTEE") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XCTE") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XCRO") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "XLAT") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "X18R") { t1template.Text = ("BUTTWLD1"); }
            if (tboxsubcom.Text == "S90R") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "SCON") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "SECC") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "SRTE") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "SRCT") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "SRCR") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "SLAR") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "SRRT") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "W90R") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "WCON") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "WECC") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "WRTE") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "WRCT") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "WRCR") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "WLAR") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "WRRT") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "X90R") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "XCON") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "XECC") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "XRTE") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "XRCT") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "XRCR") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "XLAR") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "XRRT") { t1template.Text = ("BUTTWLD2"); }
            if (tboxsubcom.Text == "RFWN") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "RFW1") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "RFWA") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "RFWB") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "FFWN") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "FFW1") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "RJWN") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "FFWA") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "FFWB") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "RFSW") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "FFSW") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "RJSW") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "LMWN") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "LFWN") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "SMWN") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "SFWN") { t1template.Text = ("FLANGE1"); }
            if (tboxsubcom.Text == "FFRW") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "RFRW") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "RRSW") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "FRSW") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "ORWN") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "OFWN") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "ORSW") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "OJWN") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "OJSW") { t1template.Text = ("FLANGE2"); }
            if (tboxsubcom.Text == "RFSO") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "FFSO") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "RJSO") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "RFBL") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "FFBL") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "BRTJ") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "RJBL") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "RFSB") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "RFBA") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "RFBB") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "FFBA") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "FFBB") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "HHBL") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "LAPJ") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "RFTD") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "FFTD") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "LAPP") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "FECO") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "MACO") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "GROV") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "GRCO") { t1template.Text = ("FLANGE3"); }
            if (tboxsubcom.Text == "RRSO") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "FRSO") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "OJTD") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "ORSO") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "OFSO") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "ORTD") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "RRTD") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "FRTD") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "RFWL") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "FFWL") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "RJWL") { t1template.Text = ("FLANGE4"); }
            if (tboxsubcom.Text == "SW90") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "TD90") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "TS90") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "SW45") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "TD45") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "SWTE") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "TDTE") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "SWCP") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "TDCP") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "SWCF") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "SWCH") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "CPST") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "TDCF") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "TDCH") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "CRSW") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "CRTD") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "LASW") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "LATD") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "UNSW") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "UNTD") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "QS90") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "QT90") { t1template.Text = ("FORGING1"); }
            if (tboxsubcom.Text == "SWRT") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "TDRT") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "STRT") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "CRST") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "SWCR") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "TDCR") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "LARS") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "LART") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "INRS") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "INST") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "HXBT") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "B1BO") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "B2BO") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "TRTR") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "SRTR") { t1template.Text = ("FORGING2"); }
            if (tboxsubcom.Text == "HXPG") { t1template.Text = ("FORGING3"); }
            if (tboxsubcom.Text == "HXPS") { t1template.Text = ("FORGING3"); }
            if (tboxsubcom.Text == "RDPG") { t1template.Text = ("FORGING3"); }
            if (tboxsubcom.Text == "RDPT") { t1template.Text = ("FORGING3"); }
            if (tboxsubcom.Text == "SQPG") { t1template.Text = ("FORGING3"); }
            if (tboxsubcom.Text == "SQPT") { t1template.Text = ("FORGING3"); }
            if (tboxsubcom.Text == "BRTD") { t1template.Text = ("FORGING3"); }
            if (tboxsubcom.Text == "NT3") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "WOL") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "WLT") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "LOB") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "VOB") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "EOB") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "SWP") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "NP3") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "NT6") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "NP4") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "NP6") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "NE3") { t1template.Text = ("OLET1"); }
            if (tboxsubcom.Text == "SOL") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "TOL") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "LOS") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "LOT") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "VOS") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "VOT") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "EOS") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "EOT") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "SLT") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "TLT") { t1template.Text = ("OLET2"); }
            if (tboxsubcom.Text == "RP6") { t1template.Text = ("OLET3"); }
            if (tboxsubcom.Text == "RP3") { t1template.Text = ("OLET3"); }
            if (tboxsubcom.Text == "RE6") { t1template.Text = ("OLET3"); }
            if (tboxsubcom.Text == "RE9") { t1template.Text = ("OLET3"); }
            if (tboxsubcom.Text == "STO") { t1template.Text = ("OLET3"); }
            if (tboxsubcom.Text == "GATS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GATT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GATF") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GAPT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EBGT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EGTD") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EGST") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EGSW") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EGTS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EGPT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EGTP") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CHCS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CHCT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CHCF") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GLOS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GLOT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GLOF") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GLOP") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GLPT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BALS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BALT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BALF") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BALP") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BAPT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CTRS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CTRT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CTRF") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BELS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BELT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BELF") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BLOS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BLOT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BLOF") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "SGCS") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "SGCT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "SGCF") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BCTT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BCSW") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "NDLT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "PSVT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "PLGV") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "RFGT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GATE") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CHCK") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GLOE") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BALL") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CTRL") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BELL") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BLOK") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "SGCK") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BLOW") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GATB") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GABT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CHCB") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "GLOB") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BALB") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BABT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "CTRB") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BELB") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "BLOB") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "SGCB") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EGBT") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "EGET") { t1template.Text = ("VALVE1"); }
            if (tboxsubcom.Text == "NIPP") { t1template.Text = ("NIPPLE1"); }
            if (tboxsubcom.Text == "CSWG") { t1template.Text = ("SWAGE1"); }
            if (tboxsubcom.Text == "ESWG") { t1template.Text = ("SWAGE1"); }
            if (tboxsubcom.Text == "YSTD") { t1template.Text = ("MISCMTL1"); }
            if (tboxsubcom.Text == "YSSW") { t1template.Text = ("MISCMTL1"); }
            if (tboxsubcom.Text == "TSBW") { t1template.Text = ("MISCMTL2"); }
            if (tboxsubcom.Text == "YSBW") { t1template.Text = ("MISCMTL2"); }
            if (tboxsubcom.Text == "BATE") { t1template.Text = ("MISCMTL2"); }
            if (tboxsubcom.Text == "BKLT") { t1template.Text = ("MISCMTL2"); }
            if (tboxsubcom.Text == "ORNG") { t1template.Text = ("MISCMTL2"); }
            if (tboxsubcom.Text == "FFFL") { t1template.Text = ("MISCMTL2"); }
            if (tboxsubcom.Text == "GHUB") { t1template.Text = ("MISCMTL2"); }
            if (tboxsubcom.Text == "PFFS") { t1template.Text = ("MISCMTL3"); }
            if (tboxsubcom.Text == "RFYS") { t1template.Text = ("MISCMTL4"); }
            if (tboxsubcom.Text == "ORNW") { t1template.Text = ("MISCMTL4"); }
            if (tboxsubcom.Text == "ORNS") { t1template.Text = ("MISCMTL4"); }
            if (tboxsubcom.Text == "COWN") { t1template.Text = ("MISCMTL4"); }
            if (tboxsubcom.Text == "GAPA") { t1template.Text = ("MISCMTL6"); }
            if (tboxsubcom.Text == "THRM") { t1template.Text = ("MISCMTL6"); }
            if (tboxsubcom.Text == "BART") { t1template.Text = ("MISCMTL8"); }




        }
        private void t1wt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }
        private void tboxcom_TextChanged(object sender, EventArgs e)
        {

        }
        private void t1template_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (t1template.Text.ToString().Contains("MISC"))
            {
                MessageBox.Show("Please Confirm the individual segments are correct");
            }
        }
        private void t1wt_TextChanged(object sender, EventArgs e)
        {

        }
        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }
        private void t2pcode_TextChanged(object sender, EventArgs e)
        { }
        private void t1desc_TextChanged(object sender, EventArgs e)
        { }
        private void t2desc_TextChanged(object sender, EventArgs e)
        { }
        private void t1mat_TextChanged(object sender, EventArgs e)
        { }
        private void t2mat_TextChanged(object sender, EventArgs e)
        { }
        private void t2desc_TextChanged_1(object sender, EventArgs e)
        { }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        { }
        private void t2gl_TextChanged(object sender, EventArgs e)
        { }
        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            DataTable im_table = itemmaster.GetDataByItemCodeLive(t1pcode.Text);

            if (im_table.Rows.Count == 0)
            {
                label1.ForeColor = Color.Red;
                label1.Text = "Status:* Null";
            }
            else
            {
                label1.ForeColor = Color.Green;
                label1.Text = "Status:* Existing";
            }
            if (t1pcode.Text == String.Empty)
            {
                label1.ForeColor = Color.Black;
                label1.Text = "Status:*";
            }

        }

        int missitemcounter = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            t1missingitems.Items.Clear();
            DataTable missing_table = missingcodes.GetData();
            this.specTableAdapter.Fill(this.pD_EDWDataSet.Spec);
            int amountofnewitems = missing_table.Rows.Count;
            int itemcounter = amountofnewitems;
            for (int i = 0; i < amountofnewitems; i++)
            {

                string thecode = missing_table.Rows[i]["ItemCode"].ToString();

                DataTable im1_table = itemmaster.GetDataByItemCodeLive(thecode);
                if (im1_table.Rows.Count == 0)
                {
                    // MessageBox.Show(thecode + im_table.Rows.Count.ToString());
                    t1missingitems.Items.Add(missing_table.Rows[i]["ItemCode"].ToString());
                }
                else
                {
                    itemcounter = (itemcounter - 1);
                    // MessageBox.Show(thecode + im_table.Rows.Count.ToString());
                    //t1missingitems.Items.Add(missing_table.Rows[i]["ItemCode"].ToString());
                }
            }
            t1missingitems.Text = "Pending PartCodes - " + itemcounter.ToString();



        }

        private void t1undoadd_Click_1(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
            DataTable im_table = itemmaster.GetDataByItemCodeLive(t1pcode.Text);
            try
            {
                tboxcom.Text = im_table.Rows[0]["COMM"].ToString();
                tboxsubcom.Text = im_table.Rows[0]["SUBCOMM"].ToString();
                tboxsize1.Text = im_table.Rows[0]["SIZE1"].ToString();
                tboxsize2.Text = im_table.Rows[0]["SIZE2"].ToString();
                tboxsch.Text = im_table.Rows[0]["SCH"].ToString();
                tboxrating.Text = im_table.Rows[0]["RATING_EC"].ToString();
                t1mat.Text = im_table.Rows[0]["MATERIAL_TYPE"].ToString();
                t1gl.Text = im_table.Rows[0]["GL_CLASS"].ToString();
                t1wt.Text = im_table.Rows[0]["WEIGHT_CONV"].ToString();
                t1sa.Text = im_table.Rows[0]["SURFACE_AREA_CONV"].ToString();
                t1desc.Text = (im_table.Rows[0]["ITEM_DESC_1"].ToString() + im_table.Rows[0]["ITEM_DESC_2"].ToString());
                t1template.Text = im_table.Rows[0]["ITEM_TEMPLATE"].ToString();
            }
            catch
            { }
        }

        private void t1imtable_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void t1missingitems_SelectedIndexChanged(object sender, EventArgs e)
        {
            label1.ForeColor = Color.Black;
            label1.Text = "Status:*";
            if (trigger == 1)
            {

                t1pcode.Text = table1.Rows[t1missingitems.SelectedIndex]["PartCode"].ToString();
                t1desc.Text = table1.Rows[t1missingitems.SelectedIndex]["Desc"].ToString();
                t1mat.Text = table1.Rows[t1missingitems.SelectedIndex]["Mat"].ToString();
                t1gl.Text = table1.Rows[t1missingitems.SelectedIndex]["GL"].ToString();
                t1wt.Text = table1.Rows[t1missingitems.SelectedIndex]["Wt"].ToString();
                t1sa.Text = table1.Rows[t1missingitems.SelectedIndex]["SA"].ToString();
                return;
            }
            else
            {
                t1desc.Text = string.Empty;
                t1mat.Text = string.Empty;
                t1gl.Text = string.Empty;
                t1wt.Text = string.Empty;
                t1sa.Text = string.Empty;

                if (t1desccheckbox.Checked == true)
                {
                    missitemcounter = t1missingitems.SelectedIndex;
                    DataTable missing_table = missingcodes.GetData();
                    string missingdesc = (missing_table.Rows[missitemcounter]["Description"].ToString());
                    t1pcode.Text = t1missingitems.SelectedItem.ToString();
                    t1desc.Text = missingdesc;
                    trigger = 0;
                }
                else
                {
                    t1pcode.Text = t1missingitems.Text;
                    DataTable im_table = itemmaster.GetDataByItemCodeLive(t1pcode.Text);
                    try
                    {
                        tboxcom.Text = im_table.Rows[0]["COMM"].ToString();
                        tboxsubcom.Text = im_table.Rows[0]["SUBCOMM"].ToString();
                        tboxsize1.Text = im_table.Rows[0]["SIZE1"].ToString();
                        tboxsize2.Text = im_table.Rows[0]["SIZE2"].ToString();
                        tboxsch.Text = im_table.Rows[0]["SCH"].ToString();
                        tboxrating.Text = im_table.Rows[0]["RATING_EC"].ToString();
                        t1mat.Text = im_table.Rows[0]["MATERIAL_TYPE"].ToString();
                        t1gl.Text = im_table.Rows[0]["GL_CLASS"].ToString();
                        t1wt.Text = im_table.Rows[0]["WEIGHT_CONV"].ToString();
                        t1sa.Text = im_table.Rows[0]["SURFACE_AREA_CONV"].ToString();
                        t1desc.Text = (im_table.Rows[0]["ITEM_DESC_1"].ToString() + im_table.Rows[0]["ITEM_DESC_2"].ToString());
                        t1template.Text = im_table.Rows[0]["ITEM_TEMPLATE"].ToString();
                    }
                    catch { }
                }
            }
            timer1.Start();
        }

        private void t1gl_TextChanged(object sender, EventArgs e)
        {

        }

        private void t2addclearcm_Opening(object sender, CancelEventArgs e)
        {

        }








        private void t3pasteandcheckbutton_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            try
            {
                // Getting Text from Clip board
                string s = Clipboard.GetText();
                //Parsing criteria: New Line 
                string[] lines = s.Split('\n');
                foreach (string ln in lines)
                {
                    listBox1.Items.Add(ln.Trim());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            listBox1.Refresh();
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                string item = (string)listBox1.Items[i];
                {
                    if (item == String.Empty)
                    {

                    }
                    else
                    {
                        DataTable im_table = itemmaster.GetDataByItemCodeLive(item);
                        if (im_table.Rows.Count == 0)
                        {
                            listBox1.Items.RemoveAt(i);
                            listBox1.Items.Insert(i, item + "  ::::  ");
                            listBox1.Refresh();
                        }
                        else
                        {
                            listBox1.Items.RemoveAt(i);
                            listBox1.Items.Insert(i, item + ":   " + im_table.Rows[0]["ITEM_DESC_1"].ToString() + im_table.Rows[0]["ITEM_DESC_2"].ToString() + ":       MAT:" + im_table.Rows[0]["MATERIAL_TYPE"].ToString() + ":  GL:" + im_table.Rows[0]["GL_CLASS"].ToString() + ":  WT:" + im_table.Rows[0]["WEIGHT_CONV"].ToString() + ":  SA:" + im_table.Rows[0]["SURFACE_AREA_CONV"].ToString());
                            listBox1.Refresh();
                        }
                        System.Threading.Thread.Sleep(50);
                    }
                }

            }
        }

        private void t3selectallandcopy_Click(object sender, EventArgs e)
        {
            try
            {
                for (int loop = 0; loop < currentlastcellt3; loop++)
                    t3breakout.Rows[loop].Selected = true;
                Clipboard.SetDataObject(
                this.t3breakout.GetClipboardContent());
            }
            catch
            {

            }

        }

        private void t3clear_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            t3breakout.Rows.Clear();
            t3breakout.Refresh();
            currentlastcellt3 = 0;

        }

        private void t3pasteandbreakout_Click(object sender, EventArgs e)
        {
            t3pasteandbreakout.Text = "Running...";
            listBox1.Items.Clear();
            try
            {
                // Getting Text from Clip board
                string s = Clipboard.GetText();
                //Parsing criteria: New Line 
                string[] lines = s.Split('\n');
                foreach (string ln in lines)
                {
                    listBox1.Items.Add(ln.Trim());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            listBox1.Refresh();
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                string item = (string)listBox1.Items[i];
                {
                    if (item == String.Empty)
                    {
                        this.t3breakout.Rows.Add();
                        t3breakout.Rows[i].Cells[0].Value = "";
                        t3breakout.Rows[i].Cells[1].Value = "";
                        t3breakout.Rows[i].Cells[2].Value = "";
                        t3breakout.Rows[i].Cells[3].Value = "";
                        t3breakout.Rows[i].Cells[4].Value = "";
                        t3breakout.Rows[i].Cells[5].Value = "";
                        t3breakout.Rows[i].Cells[6].Value = "";
                        t3breakout.Rows[i].Cells[7].Value = "";
                        currentlastcellt3++;
                    }
                    else
                    {
                        DataTable im_table = itemmaster.GetDataByItemCodeLive(item);
                        if (im_table.Rows.Count == 0)
                        {
                            //listBox1.Refresh();
                            //t3breakout.Refresh();
                            listBox1.Items.RemoveAt(i);
                            listBox1.Items.Insert(i, item + "  ::::  ");

                            this.t3breakout.Rows.Add();
                            t3breakout.Rows[i].Cells[0].Value = "";
                            t3breakout.Rows[i].Cells[1].Value = "";
                            t3breakout.Rows[i].Cells[2].Value = "";
                            t3breakout.Rows[i].Cells[3].Value = "";
                            t3breakout.Rows[i].Cells[4].Value = "";
                            t3breakout.Rows[i].Cells[5].Value = "";
                            t3breakout.Rows[i].Cells[6].Value = "";
                            t3breakout.Rows[i].Cells[7].Value = "";
                            currentlastcellt3++;
                        }
                        else
                        {
                            //listBox1.Refresh();
                            //t3breakout.Refresh();
                            listBox1.Items.RemoveAt(i);
                            listBox1.Items.Insert(i, item + "  ::::  " + im_table.Rows[0]["ITEM_DESC_1"].ToString() + im_table.Rows[0]["ITEM_DESC_2"].ToString());


                            this.t3breakout.Rows.Add();
                            t3breakout.Rows[i].Cells[0].Value = im_table.Rows[0]["SUBCOMM_DESC"].ToString();
                            t3breakout.Rows[i].Cells[1].Value = im_table.Rows[0]["SIZE1_DESC"].ToString();
                            t3breakout.Rows[i].Cells[2].Value = im_table.Rows[0]["SIZE2_DESC"].ToString();
                            t3breakout.Rows[i].Cells[3].Value = im_table.Rows[0]["SCH_DESC"].ToString();
                            t3breakout.Rows[i].Cells[4].Value = im_table.Rows[0]["RATING_EC_DESC"].ToString();
                            t3breakout.Rows[i].Cells[5].Value = im_table.Rows[0]["SGC_DESC"].ToString();
                            t3breakout.Rows[i].Cells[6].Value = im_table.Rows[0]["ITEM_DESC_1"].ToString() + im_table.Rows[0]["ITEM_DESC_2"].ToString();
                            t3breakout.Rows[i].Cells[7].Value = im_table.Rows[0]["LONG_ITEM"].ToString();
                            currentlastcellt3++;

                        }

                        System.Threading.Thread.Sleep(50);
                    }

                }

            }
            t3pasteandbreakout.Text = "Paste and Breakout";
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void t1importcm_Opening(object sender, CancelEventArgs e)
        {

        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {

        }

        private void importFromClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            t1missingitems.Items.Clear();
            try
            {
                // Getting Text from Clip board
                string s = Clipboard.GetText();
                //Parsing criteria: New Line 
                string[] lines = s.Split('\n');

                foreach (string ln in lines)
                {

                    t1missingitems.Items.Add(ln.Trim());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void toolStripMenuItem3_Click_1(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            try
            {

                for (int i = 0; i < t1missingitems.Items.Count; i++)
                {
                    listBox1.Items.Add(t1missingitems.Items[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            listBox1.Refresh();
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                string item = (string)listBox1.Items[i];
                {
                    if (item == String.Empty)
                    {

                    }
                    else
                    {
                        DataTable im_table = itemmaster.GetDataByItemCodeLive(item);
                        if (im_table.Rows.Count == 0)
                        {
                            listBox1.Items.RemoveAt(i);
                            listBox1.Items.Insert(i, item + "......:--------------------------------------------");
                            listBox1.Refresh();
                        }
                        else
                        {
                            listBox1.Items.RemoveAt(i);
                            listBox1.Items.Insert(i, item + "......:  " + im_table.Rows[0]["ITEM_DESC_1"].ToString() + im_table.Rows[0]["ITEM_DESC_2"].ToString());
                            listBox1.Refresh();
                        }
                        System.Threading.Thread.Sleep(50);
                    }
                }

            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            try
            {

                for (int i = 0; i < t1missingitems.Items.Count; i++)
                {
                    listBox1.Items.Add(t1missingitems.Items[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            listBox1.Refresh();
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                string item = (string)listBox1.Items[i];
                {
                    if (item == String.Empty)
                    {

                    }
                    else
                    {
                        DataTable im_table = itemmaster.GetDataByItemCodeLive(item);
                        if (im_table.Rows.Count == 0)
                        {
                            //listBox1.Refresh();
                            //t3breakout.Refresh();
                            listBox1.Items.RemoveAt(i);
                            listBox1.Items.Insert(i, item + ";\t" + "......:--------------------------------------------");

                            this.t3breakout.Rows.Add();
                            t3breakout.Rows[i].Cells[0].Value = "";
                            t3breakout.Rows[i].Cells[1].Value = "";
                            t3breakout.Rows[i].Cells[2].Value = "";
                            t3breakout.Rows[i].Cells[3].Value = "";
                            t3breakout.Rows[i].Cells[4].Value = "";
                            t3breakout.Rows[i].Cells[5].Value = "";
                            currentlastcellt3++;
                        }
                        else
                        {
                            //listBox1.Refresh();
                            //t3breakout.Refresh();
                            listBox1.Items.RemoveAt(i);
                            listBox1.Items.Insert(i, item + "  ::::  " + im_table.Rows[0]["ITEM_DESC_1"].ToString() + im_table.Rows[0]["ITEM_DESC_2"].ToString());


                            this.t3breakout.Rows.Add();
                            t3breakout.Rows[i].Cells[0].Value = im_table.Rows[0]["SUBCOMM_DESC"].ToString();
                            t3breakout.Rows[i].Cells[1].Value = im_table.Rows[0]["SIZE1_DESC"].ToString();
                            t3breakout.Rows[i].Cells[2].Value = im_table.Rows[0]["SIZE2_DESC"].ToString();
                            t3breakout.Rows[i].Cells[3].Value = im_table.Rows[0]["SCH_DESC"].ToString();
                            t3breakout.Rows[i].Cells[4].Value = im_table.Rows[0]["RATING_EC_DESC"].ToString();
                            t3breakout.Rows[i].Cells[5].Value = im_table.Rows[0]["SGC_DESC"].ToString();
                            currentlastcellt3++;

                        }
                        System.Threading.Thread.Sleep(50);
                    }
                }
            }
            t3pasteandbreakout.Text = "Paste and Breakout";
        }

        public void t2importfromclip_Click(object sender, EventArgs e)
        {
            triggerfile = 0;
            t2listfromclip.Items.Clear();
            try
            {
                // Getting Text from Clip board
                string s = Clipboard.GetText();
                //Parsing criteria: New Line 
                string[] lines = s.Split('\n');

                foreach (string ln in lines)
                {

                    t2listfromclip.Items.Add(ln.Trim());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void t2listfromclip_SelectedIndexChanged(object sender, EventArgs e)
        {
            string mat_string = "";
            t2gl.Text = "";
            string desctospcode = "";
            if (t2listfromclip.SelectedItem.ToString().Contains("$"))
            {
                string t2glclass4 = "";
                string v = t2listfromclip.SelectedItem.ToString();
                string[] breakout = v.Split('$');
                t2desc.Text = breakout[1].ToString();
                desctospcode = t2desc.Text;
                desctospcode = desctospcode.Replace(" ", "").Replace("-", "").Replace(".", "").Replace("*", "").Replace("/", "").Replace("\\", "").Replace("#", "").Replace("(", "").Replace(")", "").Replace("'", "").Replace("\"", "").Replace(";", "").Replace(":", "");
                t2pcode.Text = ("S" + desctospcode);
                t2desc.Text = breakout[1].ToString();
                t2mat.Text = breakout[2].ToString();
                FractionalNumber glcc = new FractionalNumber(breakout[3].ToString().Replace("\"", "").Replace("-", " "));
                double suptsize = glcc;
                if (suptsize <= 2)
                {
                    t2glclass4 = "1";
                }
                if (suptsize >= 2.5 & suptsize <= 3)
                {
                    t2glclass4 = "2";
                }
                if (suptsize >= 4 & suptsize <= 12)
                {
                    t2glclass4 = "3";
                }
                if (suptsize >= 14 & suptsize <= 16)
                {
                    t2glclass4 = "4";
                }
                if (suptsize >= 18 & suptsize <= 24)
                {
                    t2glclass4 = "5";
                }
                if (suptsize >= 26 & suptsize <= 48)
                {
                    t2glclass4 = "6";
                }
                if (suptsize > 48)
                {
                    t2glclass4 = "7";
                }

                if (t2mat.Text == String.Empty)
                {
                    t2gl.Text = "40";
                }
                if (t2mat.Text == "40")
                {
                    t2gl.Text = "403";
                }
                if (t2mat.Text == "00")
                {
                    t2gl.Text = "401";
                }
                if (t2mat.Text == "42")
                {
                    t2gl.Text = "403";
                }
                if (t2mat.Text == "60")
                {
                    t2gl.Text = "401";
                }
                if (t2mat.Text == "88")
                {
                    t2gl.Text = "404";
                }
                if (t2mat.Text == "83")
                {
                    t2gl.Text = "404";
                }
                if (t2mat.Text == "70")
                {
                    t2gl.Text = "408";
                }
                t2gl.SelectionStart = 0;
                t2gl.SelectionLength = t2gl.Text.Length;

                t2gl.Text = t2gl.Text + t2glclass4;
                t2sa.Text = breakout[4].ToString();
                t2wt.Text = breakout[5].ToString();
                t2qtytextbox.Text = breakout[6].ToString();
                mat_string = breakout[7].ToString();
            }
            else
                try
                {
                    t2desc.Text = t2listfromclip.SelectedItem.ToString();
                }
                catch
                { }
            desctospcode = t2desc.Text;
            desctospcode = desctospcode.Replace(" ", "").Replace("-", "").Replace(".", "").Replace("*", "").Replace("/", "").Replace("\\", "").Replace("#", "").Replace("(", "").Replace(")", "").Replace("'", "").Replace("\"", "").Replace(";", "").Replace(":", "");
            t2pcode.Text = ("S" + desctospcode);
            try
            {
                t2mat.Text = table2.Rows[t2listfromclip.SelectedIndex]["Mat"].ToString();
                t2gl.Text = table2.Rows[t2listfromclip.SelectedIndex]["GL"].ToString();
                t2wt.Text = table2.Rows[t2listfromclip.SelectedIndex]["Wt"].ToString();
                t2sa.Text = table2.Rows[t2listfromclip.SelectedIndex]["SA"].ToString();
            }
            catch
            { }
            if (t2mat.Text == "NA")
            {
                MessageBox.Show(mat_string);
            }
            timer4.Start();
        }






        private void t2copyasreq_Click(object sender, EventArgs e)
        { }

        private void copyAsReqToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.dataGridView3.DefaultCellStyle.WrapMode =
                DataGridViewTriState.False;
            try
            {
                for (int loop = 0; loop < dataGridView3.Rows.Count; loop++)
                    dataGridView3.Rows[loop].Selected = true;
                Clipboard.SetDataObject(
                this.dataGridView3.GetClipboardContent());
            }
            catch
            { }
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            this.dataGridView1.DefaultCellStyle.WrapMode =
            DataGridViewTriState.False;
            dataGridView1.ClearSelection();
            this.dataGridView1.Sort(this.dataGridView1.Columns["t1itemorbranch"], ListSortDirection.Ascending);
            try
            {
                for (int loop = 0; loop < currentlastcellt1 / 5; loop++)
                    dataGridView1.Rows[loop].Cells[4].Selected = true;
                Clipboard.SetDataObject(
                this.dataGridView1.GetClipboardContent());
            }
            catch
            { }
        }

        private void t1mat_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                t1gl.Text = String.Empty;
                int threedigitsize = Convert.ToInt32(tboxsize1.Text);
                double size1trimed = Convertsize1(threedigitsize);

                if (size1trimed <= 2)
                {
                    glclass4 = "1";
                }
                if (size1trimed >= 2.5 & size1trimed <= 3)
                {
                    glclass4 = "2";
                }
                if (size1trimed >= 4 & size1trimed <= 12)
                {
                    glclass4 = "3";
                }
                if (size1trimed >= 14 & size1trimed <= 16)
                {
                    glclass4 = "4";
                }
                if (size1trimed >= 18 & size1trimed <= 24)
                {
                    glclass4 = "5";
                }
                if (size1trimed >= 26 & size1trimed <= 48)
                {
                    glclass4 = "6";
                }
                if (size1trimed > 48)
                {
                    glclass4 = "7";
                }



                if (t1mat.Text == "40")
                {
                    t1gl.Text = "403";
                }
                if (t1mat.Text == "00")
                {
                    t1gl.Text = "401";
                }
                if (t1mat.Text == "42")
                {
                    t1gl.Text = "403";
                }
                if (t1mat.Text == "60")
                {
                    t1gl.Text = "401";
                }
                if (t1mat.Text == "88")
                {
                    t1gl.Text = "404";
                }
                if (t1mat.Text == "83")
                {
                    t1gl.Text = "404";
                }
                if (t1mat.Text == "84")
                {
                    t1gl.Text = "404";
                }
                if (t1mat.Text == "70")
                {
                    t1gl.Text = "408";
                }
                if (t1mat.Text == "81")
                {
                    t1gl.Text = "408";
                }

                if (t1mat.Text == String.Empty)
                {
                    t1gl.Text = "40 " + glclass4;
                }
                // else
                //  {
                //     t1gl.Text = t1gl.Text + glclass4;
                //  }
            }
            catch
            { }




        }
        string glclass4 = "";
        private object bcodeslist;

        private void tboxsize1_TextChanged(object sender, EventArgs e)
        {
            timer2.Stop();
            timer2.Start();

        }


        static public double Convertsize1(int x)
        {
            if (x == 025) { double result = 0.25; return result; }
            if (x == 050) { double result = 0.50; return result; }
            if (x == 075) { double result = 0.75; return result; }
            if (x == 001) { double result = 1; return result; }
            if (x == 125) { double result = 1.25; return result; }
            if (x == 150) { double result = 1.50; return result; }
            if (x == 175) { double result = 1.75; return result; }
            if (x == 002) { double result = 2; return result; }
            if (x == 250) { double result = 2.50; return result; }
            if (x == 275) { double result = 2.75; return result; }
            if (x == 003) { double result = 3; return result; }
            if (x == 350) { double result = 3.5; return result; }
            if (x == 004) { double result = 4; return result; }
            if (x == 450) { double result = 4.5; return result; }
            if (x == 005) { double result = 5; return result; }
            if (x == 006) { double result = 6; return result; }
            if (x == 008) { double result = 8; return result; }
            if (x == 010) { double result = 10; return result; }
            if (x == 012) { double result = 12; return result; }
            if (x == 014) { double result = 14; return result; }
            if (x == 016) { double result = 16; return result; }
            if (x == 018) { double result = 18; return result; }
            if (x == 020) { double result = 20; return result; }
            if (x == 022) { double result = 22; return result; }
            if (x == 024) { double result = 24; return result; }
            if (x == 026) { double result = 26; return result; }
            if (x == 028) { double result = 28; return result; }
            if (x == 030) { double result = 30; return result; }
            if (x == 032) { double result = 32; return result; }
            if (x == 034) { double result = 34; return result; }
            if (x == 036) { double result = 36; return result; }
            if (x == 038) { double result = 38; return result; }
            if (x == 040) { double result = 40; return result; }
            if (x == 042) { double result = 42; return result; }
            if (x == 044) { double result = 44; return result; }
            if (x == 046) { double result = 46; return result; }
            if (x == 048) { double result = 48; return result; }
            if (x == 050) { double result = 50; return result; }
            if (x == 052) { double result = 52; return result; }
            if (x == 054) { double result = 54; return result; }
            if (x == 056) { double result = 56; return result; }
            if (x == 058) { double result = 58; return result; }
            if (x == 060) { double result = 60; return result; }
            if (x == 062) { double result = 62; return result; }
            if (x == 064) { double result = 64; return result; }
            if (x == 066) { double result = 66; return result; }
            if (x == 068) { double result = 68; return result; }
            if (x == 070) { double result = 70; return result; }
            if (x == 072) { double result = 72; return result; }
            if (x == 074) { double result = 74; return result; }
            if (x == 076) { double result = 76; return result; }
            if (x == 078) { double result = 78; return result; }
            if (x == 080) { double result = 80; return result; }



            return 0;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            try
            {

                t1gl.Text = String.Empty;
                timer2.Stop();
                int threedigitsize = Convert.ToInt32(tboxsize1.Text);
                double size1trimed = Convertsize1(threedigitsize);
                if (size1trimed <= 2)
                {
                    glclass4 = "1";
                }
                if (size1trimed >= 2.5 & size1trimed <= 3)
                {
                    glclass4 = "2";
                }
                if (size1trimed >= 4 & size1trimed <= 12)
                {
                    glclass4 = "3";
                }
                if (size1trimed >= 14 & size1trimed <= 16)
                {
                    glclass4 = "4";
                }
                if (size1trimed >= 18 & size1trimed <= 24)
                {
                    glclass4 = "5";
                }
                if (size1trimed >= 26 & size1trimed <= 48)
                {
                    glclass4 = "6";
                }
                if (size1trimed > 48)
                {
                    glclass4 = "7";
                }



                if (t1mat.Text == "40")
                {
                    t1gl.Text = "403";
                }
                if (t1mat.Text == "00")
                {
                    t1gl.Text = "401";
                }
                if (t1mat.Text == "42")
                {
                    t1gl.Text = "403";
                }
                if (t1mat.Text == "60")
                {
                    t1gl.Text = "401";
                }
                if (t1mat.Text == "88")
                {
                    t1gl.Text = "404";
                }
                if (t1mat.Text == "83")
                {
                    t1gl.Text = "404";
                }
                if (t1mat.Text == "84")
                {
                    t1gl.Text = "404";
                }
                if (t1mat.Text == "70")
                {
                    t1gl.Text = "408";
                }
                if (t1mat.Text == "81")
                {
                    t1gl.Text = "408";
                }
                if (t1mat.Text == String.Empty)
                {
                    t1gl.Text = "40 " + glclass4;
                }
                else
                {
                    t1gl.Text = t1gl.Text + glclass4;
                }
            }
            catch
            { }

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void t2reqcheckbox_CheckedChanged(object sender, EventArgs e)
        {
            Form f = new Form1();
            if (t2reqcheckbox.Checked)
            {
                t2qtylabel.Visible = true;
                t2qtytextbox.Visible = true;
                t2jobnumberlabel.Visible = true;
                t2jobnumbertextbox.Visible = true;
                t2reflabel.Visible = true;
                t2reftextbox.Visible = true;
                dataGridView3.Visible = true;
                // dataGridView2.Width = tabPage2.Width - 20;
                //double dynamicheight = (tabPage2.Height / 1.32) - dataGridView2.Top;
                // dataGridView2.Height = Convert.ToInt32(dynamicheight);
                //dataGridView2.Height = (tabPage2.Height - dataGridView3.Top) - 50;
                dataGridView2.Height = dataGridView2.Bottom - dataGridView3.Location.Y;
                //dataGridView2.Top = 237;
                // dataGridView2.Height = (f.Height - 237 - dataGridView3.Height);
                // dataGridView2.Width = dataGridView3.Width - 1;



            }
            else
            {
                t2qtylabel.Visible = false;
                t2qtytextbox.Visible = false;
                t2jobnumberlabel.Visible = false;
                t2jobnumbertextbox.Visible = false;
                t2reflabel.Visible = false;
                t2reftextbox.Visible = false;
                // dataGridView2.Top = 237;
                // dataGridView2.Height = f.Bottom - 3;
                //  dataGridView2.Width = dataGridView3.Width - 1;
                //dataGridView2.Width = tabPage2.Width - 20;
                //dataGridView2.Margin.Bottom = 3;
                dataGridView2.Height = (tabPage2.Height - 241) - 3;
                dataGridView3.Visible = false;
            }

        }

        private void tboxmat_TextChanged(object sender, EventArgs e)
        {
            //timer3.Stop();
            //timer3.Start();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            //------tooooo sloooowww - use backgroundworker / threading
            // timer3.Stop();
            //  DataTable im_table = itemmaster.GetDataByMAT(tboxmat.Text);
            // if (im_table.Rows.Count == 0)
            //   {
            //}
            //  else
            // {
            //      string MATDESC = im_table.Rows[1]["SGC_DESC"].ToString();
            //       this.toolTip1.SetToolTip(tboxmat, MATDESC);
            //   }
            //   if (tboxmat.Text == String.Empty)
            //   {
            //       label1.ForeColor = Color.Black;
            //       label1.Text = "Status:*";
            //   }
        }

        private void importFromClipboardSIDMGLWtSAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            trigger = 1;
            t1missingitems.Items.Clear();
            table1.Columns.Clear();
            try
            {
                try
                {
                    table1.Columns.Add("Partcode");
                    table1.Columns.Add("Desc");
                    table1.Columns.Add("Mat");
                    table1.Columns.Add("GL");
                    table1.Columns.Add("Wt");
                    table1.Columns.Add("SA");
                }
                catch { }
                string s = Clipboard.GetText();
                int numLines = s.Split('\n').Length;
                string[] split = s.Split('\n');
                for (int j = 0; j < numLines - 1; j++)
                {
                    string heey = split[j].Replace('\t', '^');
                    //MessageBox.Show(heey);
                    string[] spliter = heey.Split('^');
                    // MessageBox.Show(spliter[0].ToString());
                    string sa_string = spliter[5].TrimEnd('\r', '\n');
                    table1.Rows.Add(new object[]{
                    spliter[0],
                    spliter[1],
                    spliter[2],
                    spliter[3],
                    spliter[4],
                    sa_string,
                    });

                }

                try
                {
                    // Getting Text from Clip board
                    string t = Clipboard.GetText();
                    //Parsing criteria: New Line 
                    string[] lines = t.Split('\n');
                    int indexer = 0;
                    foreach (DataRow rower in table1.Rows)
                    {

                        t1missingitems.Items.Add(table1.Rows[indexer]["Partcode"].ToString());
                        indexer++;
                    }
                }
                catch
                { }


            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //t1missingitems.Text = t1missingitems.
        }

        private void t4importbutton_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                System.Text.StringBuilder copy_buffer = new System.Text.StringBuilder();
                foreach (object item in listBox1.SelectedItems)
                    copy_buffer.AppendLine(item.ToString());
                if (copy_buffer.Length > 0)
                    Clipboard.SetText(copy_buffer.ToString());
            }

            if (e.Control && e.KeyCode == Keys.A)
            {
                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    listBox1.SetSelected(i, true);
                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                listBox1.SetSelected(i, true);
            }

            System.Text.StringBuilder copy_buffer = new System.Text.StringBuilder();
            foreach (object item in listBox1.SelectedItems)
                copy_buffer.AppendLine(item.ToString());
            if (copy_buffer.Length > 0)
                Clipboard.SetText(copy_buffer.ToString());

        }

        private void t3copyjdedesc_Opening(object sender, CancelEventArgs e)
        {

        }

        public void importFromFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int triggerfile = 1;
            t2listfromclip.Items.Clear();


            DataTable supportinfo = new DataTable();
            supportinfo.Columns.Add("Desc");
            supportinfo.Columns.Add("Mat");
            supportinfo.Columns.Add("GL");
            supportinfo.Columns.Add("SA");
            supportinfo.Columns.Add("Weight");
            supportinfo.Columns.Add("Qty");
            supportinfo.Columns.Add("Mat_Long");
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx;*.xls;";
            ofd.Title = "Select Excel Transmittal File";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (var stream = ofd.OpenFile())
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {

                            }
                        }
                        while (reader.NextResult());
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        reader.Close();

                        DataTable sti = result.Tables[1];

                        //DataView view = new DataView(result.Tables[1]);
                        //DataTable sti = view.ToTable(true,"Epic_Tag_No", "N.S. (Inch)");
                        foreach (DataRow row in sti.Rows)
                        {
                            object value = row["Epic_Tag_No"];
                            if (value == DBNull.Value)
                            {
                            }
                            else
                            {
                                DataRow[] rows = sti.Select("Epic_Tag_No = '" + row["Epic_Tag_No"] + "'");
                                DataRow[] check = supportinfo.Select("Desc = '" + row["Epic_Tag_No"] + "'");
                                if (check.Length != 0)
                                {

                                }
                                else
                                {
                                    try
                                    {
                                        var sumqty = rows.Sum(row2 => Convert.ToInt16(row2["QTY (Nos)"]));
                                        supportinfo.Rows.Add(new Object[] { row["Epic_Tag_No"].ToString(), "", row["N.S. (Inch)"].ToString(), "", "", sumqty });
                                    }
                                    catch
                                    {

                                    }
                                }
                            }

                        }
                    }
                }
            }

            OpenFileDialog ofd2 = new OpenFileDialog();
            ofd2.Filter = "Excel Files|*.csv;";
            ofd2.Title = "Select CSV BOM File";

            if (ofd2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (var stream = ofd2.OpenFile())
                {
                    using (var reader2 = ExcelReaderFactory.CreateCsvReader(stream))
                    {
                        do
                        {
                            while (reader2.Read())
                            {

                            }
                        }
                        while (reader2.NextResult());
                        var result2 = reader2.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        reader2.Close();
                        DataTable sawi = result2.Tables[0];
                        sawi.Columns[11].ColumnName = "Tag";
                        foreach (DataRow row in supportinfo.Rows)
                        {
                            DataRow[] rows = sawi.Select("Tag = '" + row["Desc"] + "'");
                            var sumsa = rows.Sum(row2 => Convert.ToDouble(row2["TOT_AREA"]));
                            var sumwt = rows.Sum(row2 => Convert.ToDouble(row2["TOT_WT"]));
                            try
                            {
                                //MessageBox.Show(rows[0][1].ToString());
                                if (rows[0][1].ToString().Contains("304"))
                                {
                                    row["Mat"] = "40";
                                }
                                else if (rows[0][1].ToString().Contains("316"))
                                {
                                    row["Mat"] = "42";
                                }
                                else if (rows[0][1].ToString().Contains("317"))
                                {
                                    row["Mat"] = "46";
                                }
                                else if (rows[0][1].ToString().Contains("A106"))
                                {
                                    row["Mat"] = "00";
                                }
                                else if (rows[0][1].ToString().Contains("A333"))
                                {
                                    row["Mat"] = "60";
                                }
                                else if (rows[0][1].ToString().Contains("B575"))
                                {
                                    row["Mat"] = "70";
                                }
                                else if (rows[0][1].ToString().Contains("A516"))
                                {
                                    row["Mat"] = "00";
                                }
                                else if (rows[0][1].ToString().Contains("A992"))
                                {
                                    row["Mat"] = "00";
                                }
                                else if (rows[0][1].ToString().Contains("A36"))
                                {
                                    row["Mat"] = "00";
                                }
                                else if (rows[0][1].ToString().Contains("321") || (rows[0][1].ToString().Contains("347")))
                                {
                                    row["Mat"] = "84";
                                }

                                else
                                {
                                    row["Mat"] = "NA";
                                    row["Mat_Long"] = rows[0][1].ToString();
                                }
                            }
                            catch
                            {

                            }
                            row["SA"] = sumsa;
                            row["Weight"] = sumwt;
                            //MessageBox.Show("Surface area: " + sumsa.ToString() + " Weight: " + sumwt.ToString());

                        }
                    }
                }
                int numberoflines = supportinfo.Rows.Count;
                for (int j = 0; j < numberoflines; j++)
                {
                    t2listfromclip.Items.Add("-" + j + "- $" + supportinfo.Rows[j]["Desc"].ToString() + "$" + supportinfo.Rows[j]["Mat"].ToString() + "$" + supportinfo.Rows[j]["GL"].ToString() + "$" + supportinfo.Rows[j]["SA"].ToString() + "$" + supportinfo.Rows[j]["Weight"].ToString() + "$" + supportinfo.Rows[j]["Qty"].ToString() + "$" + supportinfo.Rows[j]["Mat_Long"].ToString());
                    string Tracker_List = supportinfo.Rows[j]["Desc"].ToString() + "," + supportinfo.Rows[j]["Mat"].ToString() + "," + supportinfo.Rows[j]["GL"].ToString() + "," + supportinfo.Rows[j]["SA"].ToString() + "," + supportinfo.Rows[j]["Weight"].ToString() + "," + supportinfo.Rows[j]["Qty"].ToString() + "," + supportinfo.Rows[j]["Mat_Long"].ToString() + "," + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    string path = "V:\\MTO\\exe tools\\Item Upload Tool\\Logs\\Master_Support_List.csv";
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(Tracker_List);
                    }



                }
            }
        }

        private void importFromClipboardDescMatGLWtSAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            table2.Columns.Clear();
            trigger = 1;
            triggerfile = 0;
            t2listfromclip.Items.Clear();
            try
            {

                table2.Columns.Add("Desc");
                table2.Columns.Add("Mat");
                table2.Columns.Add("GL");
                table2.Columns.Add("Wt");
                table2.Columns.Add("SA");
                string s = Clipboard.GetText();
                int numLines = s.Split('\n').Length;
                string[] split = s.Split('\n');
                for (int j = 0; j < numLines - 1; j++)
                {
                    string tabtodash = split[j].Replace('\t', '-');
                    string[] spliter = tabtodash.Split('-');
                    table2.Rows.Add(new object[]{
                    spliter[0],
                    spliter[1],
                    spliter[2],
                    spliter[3],
                    spliter[4],
                    });

                }

                try
                {
                    // Getting Text from Clip board
                    string t = Clipboard.GetText();
                    //Parsing criteria: New Line 
                    string[] lines = t.Split('\n');
                    int indexer = 0;
                    foreach (DataRow rower in table2.Rows)
                    {

                        t2listfromclip.Items.Add(table2.Rows[indexer]["Desc"].ToString());
                        indexer++;
                    }
                }
                catch
                { }


            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //t1missingitems.Text = t1missingitems.
        }

        private void label2_Click_1(object sender, EventArgs e)
        {
            DataTable im_table = itemmaster.GetDataByItemCodeLive(t2pcode.Text);
            try
            {
                t2mat.Text = im_table.Rows[0]["MATERIAL_TYPE"].ToString();
                t2gl.Text = im_table.Rows[0]["GL_CLASS"].ToString();
                t2wt.Text = im_table.Rows[0]["WEIGHT_CONV"].ToString();
                t2sa.Text = im_table.Rows[0]["SURFACE_AREA_CONV"].ToString();
                t2desc.Text = (im_table.Rows[0]["ITEM_DESC_1"].ToString() + im_table.Rows[0]["ITEM_DESC_2"].ToString());
            }
            catch
            { }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void t2pcode_TextChanged_1(object sender, EventArgs e)
        {
            label2.ForeColor = Color.Black;
            label2.Text = "Status:*";
            if (t2pcode.Text == String.Empty)
            {
                label2.ForeColor = Color.Black;
                label2.Text = "Status:*";
            }
            timer4.Stop();
            timer4.Start();
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            timer4.Stop();
            DataTable im_table = itemmaster.GetDataByItemCodeLive(t2pcode.Text);

            if (im_table.Rows.Count == 0)
            {
                label2.ForeColor = Color.Red;
                label2.Text = "Status:* Null";
            }
            else
            {
                label2.ForeColor = Color.Green;
                label2.Text = "Status:* Existing";
            }
            if (t2pcode.Text == String.Empty)
            {
                label2.ForeColor = Color.Black;
                label2.Text = "Status:*";
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void t2mat_Enter(object sender, EventArgs e)
        {
            t2mat.SelectionStart = 0;
            t2mat.SelectionLength = t2mat.Text.Length;
        }

        private void t2gl_Enter(object sender, EventArgs e)
        {
            t2gl.SelectionStart = 0;
            t2gl.SelectionLength = t2gl.Text.Length;
        }

        private void t2wt_Enter(object sender, EventArgs e)
        {
            t2wt.SelectionStart = 0;
            t2wt.SelectionLength = t2wt.Text.Length;
        }

        private void t2sa_Enter(object sender, EventArgs e)
        {
            t2sa.SelectionStart = 0;
            t2sa.SelectionLength = t2sa.Text.Length;
        }

        private void t2qtytextbox_Enter(object sender, EventArgs e)
        {
            t2qtytextbox.SelectionStart = 0;
            t2qtytextbox.SelectionLength = t2qtytextbox.Text.Length;
        }

        private void t1mat_Enter(object sender, EventArgs e)
        {
            t1mat.SelectionStart = 0;
            t1mat.SelectionLength = t1mat.Text.Length;
        }

        private void t1gl_Enter(object sender, EventArgs e)
        {
            t1gl.SelectionStart = 0;
            t1gl.SelectionLength = t1gl.Text.Length;
        }

        private void t1wt_Enter(object sender, EventArgs e)
        {
            t1wt.SelectionStart = 0;
            t1wt.SelectionLength = t1wt.Text.Length;
        }

        private void t1sa_Enter(object sender, EventArgs e)
        {
            t1sa.SelectionStart = 0;
            t1sa.SelectionLength = t1sa.Text.Length;
        }

        private void t1pcode_Enter(object sender, EventArgs e)
        {
            t1pcode.SelectionStart = 0;
            t1pcode.SelectionLength = t1pcode.Text.Length;
        }

        private void t1desc_Enter(object sender, EventArgs e)
        {
            t1desc.SelectionStart = 0;
            t1desc.SelectionLength = t1desc.Text.Length;
        }

        private void t2desc_Enter(object sender, EventArgs e)
        {
            t2desc.SelectionStart = 0;
            t2desc.SelectionLength = t2desc.Text.Length;
        }

        private void t2pcode_Enter(object sender, EventArgs e)
        {
            t2pcode.SelectionStart = 0;
            t2pcode.SelectionLength = t2pcode.Text.Length;
        }

        private void t2importsupt_Opening(object sender, CancelEventArgs e)
        {

        }

        private void jDEItemMasterBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void t3breakout_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void t3importfromimport_Opening(object sender, CancelEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            foreach (DataGridViewRow row in t3breakout.Rows)
            {
                string[] scoms = row.Cells[0].Value.ToString().Split(null);
                foreach (string scom in scoms)
                {
                    if (!row.Cells[6].Value.ToString().ToLower().Contains(scom.ToLower()))
                        row.Cells["t3subcom"].Style.ForeColor = Color.Red;
                }

                if (!row.Cells[6].Value.ToString().ToLower().Contains(row.Cells[1].Value.ToString().ToLower() + " "))
                    row.Cells["t3size1"].Style.ForeColor = Color.Red;

                if (!row.Cells[6].Value.ToString().ToLower().Contains(" " + row.Cells[2].Value.ToString().ToLower() + " "))
                    row.Cells["t3size2"].Style.ForeColor = Color.Red;

                if (!row.Cells[6].Value.ToString().ToLower().Contains(" " + row.Cells[3].Value.ToString().ToLower() + " "))
                    row.Cells["t3sch"].Style.ForeColor = Color.Red;

                if (!row.Cells[6].Value.ToString().ToLower().Contains(" " + row.Cells[4].Value.ToString().ToLower() + " "))
                    row.Cells["t3rating"].Style.ForeColor = Color.Red;

                if (!row.Cells[6].Value.ToString().ToLower().Contains(row.Cells[5].Value.ToString().ToLower()))
                    row.Cells["t3sgc"].Style.ForeColor = Color.Red;

            }
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
           
            t4dgv.Rows.Clear();
            string right = t4jobtb.Text.Right(5).PadLeft(5,'0');
            DataTable Refdwg = reffinder.GetData(right,  t4spooltb.Text.PadLeft(6,'0'));
            if (Refdwg.Rows.Count > 0)
            {
                t4refdwgtb.Text = Refdwg.Rows[0]["spool_refDwg"].ToString();

                if (t4rb1.Checked == true)
                {
                   DataTable bom_Table = bomfinder.GetDataBy(t4refdwgtb.Text + "%");
                    if (bom_Table.Rows.Count > 0)
                    {

                        for (int i = 0; i < bom_Table.Rows.Count; i++)
                        {
                            this.t4dgv.Rows.Add();
                            t4dgv.Rows[i].Cells[0].Value = bom_Table.Rows[i]["Source"].ToString();
                            t4dgv.Rows[i].Cells[1].Value = bom_Table.Rows[i]["Pipeline_Reference"].ToString();
                            t4dgv.Rows[i].Cells[2].Value = bom_Table.Rows[i]["Piping_Spec"].ToString();
                            t4dgv.Rows[i].Cells[3].Value = bom_Table.Rows[i]["Item_Code"].ToString();
                            t4dgv.Rows[i].Cells[4].Value = bom_Table.Rows[i]["Description"].ToString();
                            t4dgv.Rows[i].Cells[5].Value = bom_Table.Rows[i]["Qty2"].ToString();
                            t4dgv.Rows[i].Cells[6].Value = bom_Table.Rows[i]["Long_ID"].ToString();
                            t4dgv.Rows[i].Cells[7].Value = bom_Table.Rows[i]["JDE_Desc"].ToString();
                        }
                        t4dgv.AutoResizeColumns();
                    }
                }
                else if (t4rb2.Checked == true)
                {
                    DataTable bom_Table = bomfinder.GetData(t4refdwgtb.Text + "%");
                    if (bom_Table.Rows.Count > 0)
                    {

                        for (int i = 0; i < bom_Table.Rows.Count; i++)
                        {
                            this.t4dgv.Rows.Add();
                            t4dgv.Rows[i].Cells[0].Value = bom_Table.Rows[i]["Source"].ToString();
                            t4dgv.Rows[i].Cells[1].Value = bom_Table.Rows[i]["Piecemark"].ToString();
                            t4dgv.Rows[i].Cells[2].Value = bom_Table.Rows[i]["Piping_Spec"].ToString();
                            t4dgv.Rows[i].Cells[3].Value = bom_Table.Rows[i]["Item_Code"].ToString();
                            t4dgv.Rows[i].Cells[4].Value = bom_Table.Rows[i]["Description"].ToString();
                            t4dgv.Rows[i].Cells[5].Value = bom_Table.Rows[i]["Qty2"].ToString();
                            t4dgv.Rows[i].Cells[6].Value = bom_Table.Rows[i]["Long_ID"].ToString();
                            t4dgv.Rows[i].Cells[7].Value = bom_Table.Rows[i]["JDE_Desc"].ToString();
                        }
                        t4dgv.AutoResizeColumns();
                    }
                }


                



            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < (t4BarcodeBomList.Rows.Count - 1); i++)
            {

                
                if (t4rb1.Checked == true)
                {
                    DataTable Refdwg = reffinder.GetDataBy(t4BarcodeBomList.Rows[i].Cells["bcodes"].FormattedValue.ToString());
                    t4refdwgtb.Text = Refdwg.Rows[0]["spool_pcmark"].ToString();
                    DataTable bom_Table = bomfinder.GetData(t4refdwgtb.Text + "%");
                    if (bom_Table.Rows.Count == 0)
                    {

                    }
                    for (int o = 0; o < bom_Table.Rows.Count; o++)
                    {
                        this.t4dgv.Rows.Add(
                        bom_Table.Rows[o]["Source"].ToString(),
                        bom_Table.Rows[o]["Pipeline_Reference"].ToString(),
                        bom_Table.Rows[o]["Piping_Spec"].ToString(),
                        bom_Table.Rows[o]["Item_Code"].ToString(),
                        bom_Table.Rows[o]["Description"].ToString(),
                        bom_Table.Rows[o]["Qty2"].ToString(),
                        bom_Table.Rows[o]["Long_ID"].ToString(),
                        bom_Table.Rows[o]["JDE_Desc"].ToString());
                    }
                }
                else if (t4rb2.Checked == true)
                {
                    DataTable Refdwg = reffinder.GetDataBy(t4BarcodeBomList.Rows[i].Cells["bcodes"].FormattedValue.ToString());
                    t4refdwgtb.Text = Refdwg.Rows[0]["spool_pcmark"].ToString();
                    DataTable bom_Table = bomfinder.GetData(t4refdwgtb.Text + "%");
                    if (bom_Table.Rows.Count == 0)
                    {

                    }
                    for (int o = 0; o < bom_Table.Rows.Count; o++)
                    {
                        this.t4dgv.Rows.Add(
                        bom_Table.Rows[o]["Source"].ToString(),
                        bom_Table.Rows[o]["Piecemark"].ToString(),
                        bom_Table.Rows[o]["Piping_Spec"].ToString(),
                        bom_Table.Rows[o]["Item_Code"].ToString(),
                        bom_Table.Rows[o]["Description"].ToString(),
                        bom_Table.Rows[o]["Qty2"].ToString(),
                        bom_Table.Rows[o]["Long_ID"].ToString(),
                        bom_Table.Rows[o]["JDE_Desc"].ToString());
                    }
                }

            }
            t4dgv.AutoResizeColumns();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            int l = 0;
            string s = Clipboard.GetText();
            string[] lines = s.Split('\n','\r');
            foreach (string ln in lines)
            {
                if (ln.Length == 6)
                {
                    t4BarcodeBomList.Rows.Add("*" + ln + "*");
                }
            }
        }

        private void t4ClearList_Click(object sender, EventArgs e)
        {
            t4BarcodeBomList.Rows.Clear();
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            t4dgv.Rows.Clear();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}


