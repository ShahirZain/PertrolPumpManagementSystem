using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Data.SqlClient;

namespace Petro_F
{
    public partial class Cashbook : Form
    {
        Form2 f2 = new Form2();

        DateTime a = DateTime.Now.Date;
        public Cashbook()
        {
            InitializeComponent();
        }

        private void Cashbook_Load(object sender, EventArgs e)
        {
            //dateTimePicker1.Format = DateTimePickerFormat.Custom;
            //dateTimePicker1.CustomFormat = "yyyy-mm-dd";
            this.textBox1.Text = dateTimePicker1.Value.Date.ToString();
            DGW();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = a.AddMonths(-1).ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var aPlusMonth = a.AddMonths(1);
            this.textBox1.Text = aPlusMonth.ToString();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }


        public void DGW() {

            try
            {
                f2.sqlConnection1.Open();
                SqlCommand cmd1 = new SqlCommand("select * from cashbook", f2.sqlConnection1);
                SqlDataReader dr1 = cmd1.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr1);
                dataGridView1.DataSource = dt;
                f2.sqlConnection1.Close();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString());
            }

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNO"].Value = (e.RowIndex + 1).ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataRow toInsert = dt.NewRow();

            // insert in the desired place
            dt.Rows.InsertAt(toInsert, this.dataGridView1.RowCount +1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataRow toInsert = dt.NewRow();

            // insert in the desired place
            dt.Rows.InsertAt(toInsert, this.dataGridView1.RowCount - 1);
        }




        public void gridtopdf(DataGridView dgw, string filename)
        {
            
            BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1250, BaseFont.EMBEDDED);
            PdfPTable pt = new PdfPTable(dgw.Columns.Count);
            pt.DefaultCell.Padding = 10;
            pt.WidthPercentage = 100;
            pt.HorizontalAlignment = Element.ALIGN_LEFT;
            pt.DefaultCell.BorderWidth = 1;
            iTextSharp.text.Font text = new iTextSharp.text.Font(bf, 10, iTextSharp.text.Font.NORMAL);
            // add header
               
            foreach (DataGridViewColumn colum in dataGridView1.Columns)
            {
                Color cr = new Color();
                PdfPCell cell = new PdfPCell(new Phrase(colum.HeaderText, text));
                cell.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);// basecolor =color
                pt.AddCell(cell);
            }
            //add data row
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                
                foreach (DataGridViewCell cell in row.Cells)
                {
                    PdfPCell cells = new PdfPCell(new Phrase(cell.Value.ToString(), text));
                    pt.AddCell(cells);

                }
            }
            var savefile = new SaveFileDialog();
            savefile.FileName = filename;

            savefile.DefaultExt = ".pdf";
            if (savefile.ShowDialog() == DialogResult.OK)
            {
                using (FileStream str = new FileStream(savefile.FileName, FileMode.Create))
                {
                    Document pdfdoc = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
                    PdfWriter.GetInstance(pdfdoc, str);
                    pdfdoc.Open();
                    pdfdoc.Add(pt);
                    pdfdoc.Close();
                    str.Close();
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            gridtopdf(dataGridView1, "");
        }

        private void button8_Click(object sender, EventArgs e)
        {
           
        }
    }

}
