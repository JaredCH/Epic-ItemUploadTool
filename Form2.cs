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
    public partial class Form2 : Form
    {
        private Form1 mOpener = null;
        public Form2()
        {
            InitializeComponent();
        }

        public Form1 Opener
        {
            set { this.mOpener = value; }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            glgrid3.Rows.Add("1", "Carbon / Gr.6 Low-Temp (Gps. 0,6)", "Gps. 0, 6");
            glgrid3.Rows.Add("2", "Chrome p11 & 4130 (Gp. 1)", "Gp. 1");
            glgrid3.Rows.Add("3", "Stainless 304 & 316 (Gp. 1", "Gp. 4");
            glgrid3.Rows.Add("4", "SS 321, 347, 410 & High Alloy (Gps. 8, 14", "Gps. 8, 14");
            glgrid3.Rows.Add("5", "Chrome P22 & P5 (Gp. 2)", "Gp. 2");
            glgrid3.Rows.Add("6", "Chrome P9 & Gr. 3 Low-Temp (Gp. 3)", "Gp. 3");
            glgrid3.Rows.Add("7", "Chrome P91 & 92 (Gps. 10, 12)", "Gps. 10, 12");
            glgrid3.Rows.Add("8", "Hastelloy, Cu, Ni, Al, Ti, Zr, (Gps. 5, 7, 9, 11, 13)", "Gps. 5, 7, 9, 11, 13");

            sagrid4.Rows.Add("1", "2\" and less");
            sagrid4.Rows.Add("2", "2.5\" to 3\"");
            sagrid4.Rows.Add("3", "4\" to 12\"");
            sagrid4.Rows.Add("4", "14\" to 16\"");
            sagrid4.Rows.Add("5", "18\" to 24\"");
            sagrid4.Rows.Add("6", "26\" to 48\"");
            sagrid4.Rows.Add("7", "Over 48\"");

        }

        private void sagrid4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {}



        /// <summary>
        /// //////////Return Selection of GL Class Code to Form 1///////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gldonebtn_Click(object sender, EventArgs e)
        {
            string gl3code ="";
            string gl4code ="";
            string tab = this.mOpener.whichtab;
            if (glgrid3.SelectedCells.Count > 0)
            {
                gl3code = glgrid3.SelectedCells[0].Value.ToString();
            }
            if (sagrid4.SelectedCells.Count > 0)
            {
                gl4code = sagrid4.SelectedCells[0].Value.ToString();
            }


                gl3code =  gl3code.Replace("Carbon / Gr.6 Low-Temp (Gps. 0,6)", "1").Replace("Chrome p11 & 4130 (Gp. 1)", "2").Replace("Stainless 304 & 316 (Gp. 1", "3").Replace("SS 321, 347, 410 & High Alloy (Gps. 8, 14", "4").Replace("Chrome P22 & P5 (Gp. 2)", "5").Replace("Chrome P9 & Gr. 3 Low-Temp (Gp. 3)", "6").Replace("Chrome P91 & 92 (Gps. 10, 12)", "7").Replace("Hastelloy, Cu, Ni, Al, Ti, Zr, (Gps. 5, 7, 9, 11, 13)", "8");
                gl4code = gl4code.Replace("2\" and less", "1").Replace("2.5\" to 3\"", "2").Replace("4\" to 12\"", "3").Replace("14\" to 16\"", "4").Replace("18\" to 24\"", "5").Replace("26\" to 48\"", "6").Replace("Over 48\"", "7");



            if (this.mOpener != null)
            {
                if (tab.Equals("t1"))
                {
                    this.mOpener.SetTextt1("40" + gl3code.ToString() + gl4code.ToString());
                }
                else
                {
                    this.mOpener.SetTextt2("40" + gl3code.ToString() + gl4code.ToString());
                }
            }
            this.Close();
        }

        private void glgrid3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
