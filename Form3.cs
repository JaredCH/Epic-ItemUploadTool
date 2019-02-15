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
    public partial class Form3 : Form
    {
        private Form1 mOpener = null;
        public Form3()
        {
            InitializeComponent();
        }

        public Form1 Opener
        {
            set { this.mOpener = value; }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            var dataIndexNo = dataGridView1.Rows[e.RowIndex].Index.ToString();
            string cellValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

            //MessageBox.Show("The row index = " + dataIndexNo.ToString() + " and the row data in second column is: "
                //+ cellValue.ToString());

                string matcode = "";
                string tab = this.mOpener.whichtab;
                if (dataGridView1.SelectedCells.Count > 0)
                {
                    matcode = cellValue.ToString();
                }


            matcode = matcode.Replace("CARBON STEEL WLD E/0.5 MOLY", "0A").Replace("316/316L ABS APPROVED","43").Replace("CS & LTCS @ >= -20F", "00").Replace("JACKETED STAINLESS STEEL","44").Replace("HI-YIELD CS", "01").Replace("SOUR SERVICE 316L STNLS STEEL","45").Replace("NACE CARBON STEEL", "02").Replace("317 STAINLESS STEEL","46").Replace("GALVANIZED CS-ABS APPROVED", "03").Replace("309/310 STAINLESS STEEL ","48").Replace("RUBBER LINED CS", "04").Replace("B88 COPPER","50").Replace("COAT & WRAP CS", "05").Replace("BRONZE","55").Replace("CEMENT LINED CS", "06").Replace("A33 GR1,6 LTCS @ <= -50F","60").Replace("CARBON STEEL-ABS APPROVED", "07").Replace("SOUR SERVICE LOW TEMP","61").Replace("JACKETED CS", "08").Replace("A33 GR1, WLD WITH CS FILLER","62").Replace("GALVANIZED CS", "09").Replace("A33 GR4&9 LOW TEMP CARBON STL","64").Replace("LW CHR", "1A").Replace("A33 GR1&6 LT CS STL ABS APPROV","65").Replace("CRB MLY", "1B").Replace("HASTELLOY","70").Replace("4130 CARBON STEEL NORMALIZED", "1C").Replace("99% NICKEL","72").Replace("P11 CHROME 1-1/ 4% CHROME", "10").Replace("TITANIUM","76").Replace("ALY4130", "11").Replace("ZIRCONIUM","77").Replace("P1 CHROME", "12").Replace("XM-11","8A").Replace("ALT 4140", "13").Replace("SUPER DUPLEX (VOID)","8B").Replace("ALTY 4130 ABS APPROVED", "14").Replace("SUPER DUPLEX ABS APPROVED","8C").Replace("ALTY 4140 ABS APPROVED", "15").Replace("ALLOY 625 ABD APPROVED","8D").Replace("P2 CHROME", "16").Replace("SUPER DUPLEX","8E").Replace("P3 CHROME", "17").Replace("INCONEL","80").Replace("912 CHROME", "18").Replace("INCOLOY","81").Replace("CS HSTR", "19").Replace("ALLOY20","82").Replace("P22 CHROME 2-1/2%", "20").Replace("MONEL","83").Replace("P5 CHORME 5% CR", "24").Replace("321/347 STAINLESS STEEL","84").Replace("P3b", "28").Replace("CUNI ABS APPROVED CUPRO-NICKEL","85").Replace("P21 CHROME", "29").Replace("CUPRO-NICKEL","86").Replace("P9  CHROME 9% CR", "30").Replace("DUPLEX ABS APPROVED","87").Replace("P7 CHROME 7% CR", "32").Replace("DUPLEX STAINLESS STEEL","88").Replace("FERRITIC CHROME", "34").Replace("AL6XN","89").Replace("A33 GR 3, 3-1/2% NI LOW TEMP", "36").Replace("ALUMINUM","90").Replace("P91 CHROME", "38").Replace("YOLOY","91").Replace("P92 CHROME", "39").Replace("4130 CS NORM - VOIDED","92").Replace("304/304L STAINLESS STEEL", "40").Replace("SHIPLOOSE SUPPORT","98").Replace("316/316L STAINLESS STEEL", "42").Replace("MISCELLANEOUS","99");



            if (this.mOpener != null)
                {
                    if (tab.Equals("t1"))
                    {
                        this.mOpener.SetTextt1m(matcode.ToString());
                    }
                    else
                    {
                        this.mOpener.SetTextt2m(matcode.ToString());
                    }
                }
                this.Close();

            }



        private void Form3_Load(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add("0A", "CARBON STEEL WLD E/0.5 MOLY", "","43","316/316L ABS APPROVED");
            dataGridView1.Rows.Add("00", "CS & LTCS @ >= -20F", "","44","JACKETED STAINLESS STEEL");
            dataGridView1.Rows.Add("01", "HI-YIELD CS", "","45","SOUR SERVICE 316L STNLS STEEL");
            dataGridView1.Rows.Add("02", "NACE CARBON STEEL", "","46","317 STAINLESS STEEL");
            dataGridView1.Rows.Add("03", "GALVANIZED CS-ABS APPROVED", "","48","309/310 STAINLESS STEEL ");
            dataGridView1.Rows.Add("04", "RUBBER LINED CS", "","50","B88 COPPER");
            dataGridView1.Rows.Add("05", "COAT & WRAP CS", "","55","BRONZE");
            dataGridView1.Rows.Add("06", "CEMENT LINED CS", "","60","A333 GR1,6 LTCS @ <= -50F");
            dataGridView1.Rows.Add("07", "CARBON STEEL-ABS APPROVED", "","61","SOUR SERVICE LOW TEMP");
            dataGridView1.Rows.Add("08", "JACKETED CS", "","62","A333 GR1, WLD WITH CS FILLER");
            dataGridView1.Rows.Add("09", "GALVANIZED CS", "","64","A333 GR4&9 LOW TEMP CARBON STL");
            dataGridView1.Rows.Add("1A", "LW CHR", "","65","A333 GR1&6 LT CS STL ABS APPROV");
            dataGridView1.Rows.Add("1B", "CRB MLY", "","70","HASTELLOY");
            dataGridView1.Rows.Add("1C", "4130 CARBON STEEL NORMALIZED", "","72","99% NICKEL");
            dataGridView1.Rows.Add("10", "P11 CHROME 1-1/ 4% CHROME", "","76","TITANIUM");
            dataGridView1.Rows.Add("11", "ALY4130", "","77","ZIRCONIUM");
            dataGridView1.Rows.Add("12", "P1 CHROME", "","8A","XM-11");
            dataGridView1.Rows.Add("13", "ALT 4140", "","8B","SUPER DUPLEX (VOID)");
            dataGridView1.Rows.Add("14", "ALTY 4130 ABS APPROVED", "","8C","SUPER DUPLEX ABS APPROVED");
            dataGridView1.Rows.Add("15", "ALTY 4140 ABS APPROVED", "","8D","ALLOY 625 ABD APPROVED");
            dataGridView1.Rows.Add("16", "P2 CHROME", "","8E","SUPER DUPLEX");
            dataGridView1.Rows.Add("17", "P3 CHROME", "","80","INCONEL");
            dataGridView1.Rows.Add("18", "912 CHROME", "","81","INCOLOY");
            dataGridView1.Rows.Add("19", "CS HSTR", "","82","ALLOY20");
            dataGridView1.Rows.Add("20", "P22 CHROME 2-1/2%", "","83","MONEL");
            dataGridView1.Rows.Add("24", "P5 CHORME 5% CR", "","84","321/347 STAINLESS STEEL");
            dataGridView1.Rows.Add("29", "P21 CHROME", "","85","CUNI ABS APPROVED CUPRO-NICKEL");
            dataGridView1.Rows.Add("30", "P9  CHROME 9% CR", "","86","CUPRO-NICKEL");
            dataGridView1.Rows.Add("32", "P7 CHROME 7% CR", "","87","DUPLEX ABS APPROVED");
            dataGridView1.Rows.Add("34", "FERRITIC CHROME", "","88","DUPLEX STAINLESS STEEL");
            dataGridView1.Rows.Add("36", "A333 GR 3, 3-1/2% NI LOW TEMP", "","89","AL6XN");
            dataGridView1.Rows.Add("38", "P91 CHROME", "","90","ALUMINUM");
            dataGridView1.Rows.Add("39", "P92 CHROME", "","91","YOLOY");
            dataGridView1.Rows.Add("40", "304/304L STAINLESS STEE", "","92","4130 CS NORM - VOIDED");
            dataGridView1.Rows.Add("42", "316/316L STAINLESS STEEL", "","98","SHIPLOOSE SUPPORT");
            dataGridView1.Rows.Add("", "", "","99","MISCELLANEOUS");

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
