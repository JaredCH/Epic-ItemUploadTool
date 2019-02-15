using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ItemUploadTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int currentlastcellt1 = 0;
        int currentlastcellt2 = 0;



        private void t1pcode_TextChanged(object sender, EventArgs e)
        {

        }
        private void t2pcode_TextChanged(object sender, EventArgs e)
        {

        }
        private void t1desc_TextChanged(object sender, EventArgs e)
        {

        }
        private void t2desc_TextChanged(object sender, EventArgs e)
        {

        }
        private void t1mat_TextChanged(object sender, EventArgs e)
        {

        }
        private void t2mat_TextChanged(object sender, EventArgs e)
        {

        }

        private void t2desc_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void t2mat_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void t2convertlbl_Click(object sender, EventArgs e)
        {
            t1pcode.Text = t1desc.Text;
        }











/// <Undo_Add>
/// ///////////////////////////////////////////////////////////////////////////////////////////////////////
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
        private void t1undoadd_Click(object sender, EventArgs e)
        {
            {
                for (int i = 0; i < 5; i++)
                {
                    dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);
                }
                currentlastcellt1 = currentlastcellt1 - 5;
            }
        }

        private void t2undoadd_Click(object sender, EventArgs e)
        {
            {
                for (int i = 0; i < 5; i++)
                {
                    dataGridView2.Rows.RemoveAt(dataGridView1.Rows.Count - 2);
                }
                currentlastcellt2 = currentlastcellt2 - 5;
            }
        }



/// <Undo_Sort>
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





        /// <Copy_Function>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t1copy_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Sort(this.dataGridView1.Columns["itemorbranch"], ListSortDirection.Ascending);
            for (int loop = 0; loop < currentlastcellt1; loop++)
                dataGridView1.Rows[loop].Selected = true;
            Clipboard.SetDataObject(
            this.dataGridView1.GetClipboardContent());
        }
        private void t2copy_Click(object sender, EventArgs e)
        {
            this.dataGridView2.Sort(this.dataGridView2.Columns["t2itemorbranch"], ListSortDirection.Ascending);
            for (int loop = 0; loop < currentlastcellt2; loop++)
                dataGridView2.Rows[loop].Selected = true;
            Clipboard.SetDataObject(
            this.dataGridView2.GetClipboardContent());
        }




/// <Clear_Data>
/// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
        private void t1cleardata_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
        }

        private void t2cleardata_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Refresh();
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
        }
        private void t2reset_Click(object sender, EventArgs e)
        {
            t2pcode.Text = string.Empty;
            t2desc.Text = string.Empty;
            t2mat.Text = string.Empty;
            t2gl.Text = string.Empty;
            t2wt.Text = string.Empty;
            t2sa.Text = string.Empty;
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void t1add_Click(object sender, EventArgs e)
        {
            string whatspcode = t1pcode.Text.Substring(0, 1);
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
                Int32 lastrow = dataGridView1.Rows.Count - 1;
                this.dataGridView1.Rows.Add();

                dataGridView1.Rows[lastrow].Cells[0].Value = "02";
                dataGridView1.Rows[lastrow].Cells[1].Value = counter;
                dataGridView1.Rows[lastrow].Cells[2].Value = item;

                dataGridView1.Rows[lastrow].Cells[4].Value = t1pcode.Text;
                dataGridView1.Rows[lastrow].Cells[5].Value = t1pcode.Text;
                dataGridView1.Rows[lastrow].Cells[6].Value = t1desc.Text;
                dataGridView1.Rows[lastrow].Cells[17].Value = t1mat.Text;
                dataGridView1.Rows[lastrow].Cells[24].Value = t1gl.Text;
                dataGridView1.Rows[lastrow].Cells[30].Value = t1wt.Text;
                dataGridView1.Rows[lastrow].Cells[31].Value = t1sa.Text;

                ///////////////////////////////////////////////////////////////////////////////////////////

                dataGridView1.Rows[lastrow].Cells[10].Value = t1pcode.Text.Substring(0, 1);
                

                if (whatspcode == ("O"))
                {
                    dataGridView1.Rows[lastrow].Cells[11].Value = t1pcode.Text.Substring(1, 3);
                    dataGridView1.Rows[lastrow].Cells[12].Value = t1pcode.Text.Substring(4, 3);
                }
                else
                {
                    dataGridView1.Rows[lastrow].Cells[11].Value = t1pcode.Text.Substring(1, 4);
                    dataGridView1.Rows[lastrow].Cells[12].Value = t1pcode.Text.Substring(5, 3);
                }




                dataGridView1.Rows[lastrow].Cells[15].Value = t1pcode.Text.Substring(t1pcode.Text.Length -3);


                dataGridView1.Rows[lastrow].Cells[9].Value = "A";
                if (whatspcode == ("P"))
                {
                    dataGridView1.Rows[lastrow].Cells[21].Value = "FT";
                }
                else
                {
                    dataGridView1.Rows[lastrow].Cells[21].Value = "EA";
                }
                dataGridView1.Rows[lastrow].Cells[22].Value = "LB";
                dataGridView1.Rows[lastrow].Cells[23].Value = "SF";
                dataGridView1.Rows[lastrow].Cells[25].Value = "2";
                dataGridView1.Rows[lastrow].Cells[26].Value = "0";
                dataGridView1.Rows[lastrow].Cells[27].Value = "9";
                dataGridView1.Rows[lastrow].Cells[28].Value = "S";
                dataGridView1.Rows[lastrow].Cells[29].Value = "3650";
                dataGridView1.Rows[lastrow].Cells[33].Value = "Y";
                i++;
            }
            currentlastcellt1 = currentlastcellt1 + 5;
        }
        /// < Add_Data-Supports>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t2add_Click(object sender, EventArgs e)
        {
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
                dataGridView2.Rows[lastrow].Cells[27].Value = "9";
                dataGridView2.Rows[lastrow].Cells[28].Value = "S";
                dataGridView2.Rows[lastrow].Cells[29].Value = "3650";
                dataGridView2.Rows[lastrow].Cells[33].Value = "Y";
                i++;
            }
            currentlastcellt2 = currentlastcellt2 + 5;
        }
        /// <Add_Data_Clear-Supports>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void t2addclear_Click(object sender, EventArgs e)
        {
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
                dataGridView2.Rows[lastrow].Cells[27].Value = "9";
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

        private void t2reset_Click_1(object sender, EventArgs e)
        {
            t2pcode.Text = string.Empty;
            t2desc.Text = string.Empty;
            t2mat.Text = string.Empty;
            t2gl.Text = string.Empty;
            t2wt.Text = string.Empty;
            t2sa.Text = string.Empty;
        }

    }
}
