using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using ExcelDataReader;
using LiveCharts.Wpf;
using LiveCharts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GraduationProject1
{
    public partial class Babel : Form
    {

        static string url = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\testdata.mdf;Integrated Security=True";
        SqlConnection conn = new SqlConnection(url);


        public Babel()
        {
            InitializeComponent();
        }

        private void Babel_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'testdataDataSet1.babel' table. You can move, or remove it, as needed.
            this.babelTableAdapter1.Fill(this.testdataDataSet1.babel);
         

        }

        int Id;
        private void DisplayData()
        {
            conn.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter adapt = new SqlDataAdapter("select * from babel", conn);
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            conn.Close();
           
        }
        private void ClearData()
        {
            fn.Text = "";
            dn.Text = "";
            gn.Text = "";
            gdn.Text = "";
            nn.Text = "";
            s.Text = "";
            mn.Text = "";
            mdn.Text = "";
            mgn.Text = "";
            d.Text = "";
            n.Text = "";
            c.Text = "";
            gc.Text = "";
            gu.Text = "";
            gcol.Text = "";
            gd.Text = "";
            gm.Text = "";
            cm.Text = "";
            gnu.Text = "";
            snn.Text = "";
            pn.Text = "";
        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {


            if (string.IsNullOrEmpty(fn.Text) || string.IsNullOrEmpty(dn.Text) || string.IsNullOrEmpty(gn.Text)
                || string.IsNullOrEmpty(gdn.Text) || string.IsNullOrEmpty(nn.Text) || string.IsNullOrEmpty(mn.Text)
                || string.IsNullOrEmpty(mdn.Text) || string.IsNullOrEmpty(mgn.Text) || string.IsNullOrEmpty(d.Text)
                || string.IsNullOrEmpty(n.Text) || string.IsNullOrEmpty(c.Text) || string.IsNullOrEmpty(gc.Text)
                 || string.IsNullOrEmpty(gu.Text) || string.IsNullOrEmpty(gcol.Text) || string.IsNullOrEmpty(gd.Text)
                  || string.IsNullOrEmpty(gm.Text) || string.IsNullOrEmpty(cm.Text) || string.IsNullOrEmpty(gnu.Text)
                   || string.IsNullOrEmpty(pn.Text) || string.IsNullOrEmpty(s.Text) || string.IsNullOrEmpty(snn.Text))
            {
                MessageBox.Show("please enter all required information");
            }
            else
            {
                SqlCommand cmd = new SqlCommand("insert into babel( Fn,Dn,Gn,Gdn,Nn,SE,Mn,Mdn,Mgn,D,N,C,Cd,Gc,Gu,Gcol,Gd,Gm,Cm,Gnu,Snn1,Nnd,Pn) values(@fn,@dn,@gn,@gdn,@nn,@se,@mn,@mdn,@mgn,@d,@n,@c,@cd,@gc,@gu,@gcol,@gd,@gm,@cm,@gnu,@snn,@nnd,@pn)", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@fn", fn.Text);
                cmd.Parameters.AddWithValue("@dn", dn.Text);
                cmd.Parameters.AddWithValue("@gn", gn.Text);
                cmd.Parameters.AddWithValue("@gdn", gdn.Text);
                cmd.Parameters.AddWithValue("@nn", nn.Text);
                cmd.Parameters.AddWithValue("@se", s.Text);
                cmd.Parameters.AddWithValue("@mn", mn.Text);
                cmd.Parameters.AddWithValue("@mdn", mdn.Text);
                cmd.Parameters.AddWithValue("@mgn", mgn.Text);
                cmd.Parameters.AddWithValue("@d", d.Text);
                cmd.Parameters.AddWithValue("@n", n.Text);
                cmd.Parameters.AddWithValue("@c", c.Text);
                cmd.Parameters.AddWithValue("@cd", cd.Value.Date);
                cmd.Parameters.AddWithValue("@gc", gc.Text);
                cmd.Parameters.AddWithValue("@gu", gu.Text);
                cmd.Parameters.AddWithValue("@gcol", gcol.Text);
                cmd.Parameters.AddWithValue("@gd", gd.Text);
                cmd.Parameters.AddWithValue("@gm", gm.Text);
                cmd.Parameters.AddWithValue("@cm", cm.Text);
                cmd.Parameters.AddWithValue("@gnu", gnu.Text);
                cmd.Parameters.AddWithValue("@snn", snn.Text);
                cmd.Parameters.AddWithValue("@nnd", nnd.Value.Date);
                cmd.Parameters.AddWithValue("@pn", pn.Text);
                cmd.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Record Inserted Successfully");
                DisplayData();
                ClearData();
            }
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            ClearData();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {

            if (Id > 0)
            {
                SqlCommand cmd = new SqlCommand("update babel set Fn=@fn,Dn=@dn,Gn=@gn,Gdn=@gdn,Nn=@nn,SE=@se ,Mn=@mn,Mdn=@mdn,Mgn=@mgn,D=@d,N=@n,C=@c,Cd=@cd,Gc=@gc,Gu=@gu,Gcol=@gcol,Gd=@gd,Gm=@gm,Cm=@cm,Gnu=@gnu,Snn1=@snn,Nnd=@nnd,Pn=@pn where Id=@id", conn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@id", this.Id);
                cmd.Parameters.AddWithValue("@fn", fn.Text);
                cmd.Parameters.AddWithValue("@dn", dn.Text);
                cmd.Parameters.AddWithValue("@gn", gn.Text);
                cmd.Parameters.AddWithValue("@gdn", gdn.Text);
                cmd.Parameters.AddWithValue("@nn", nn.Text);
                cmd.Parameters.AddWithValue("@se", s.Text);
                cmd.Parameters.AddWithValue("@mn", mn.Text);
                cmd.Parameters.AddWithValue("@mdn", mdn.Text);
                cmd.Parameters.AddWithValue("@mgn", mgn.Text);
                cmd.Parameters.AddWithValue("@d", d.Text);
                cmd.Parameters.AddWithValue("@n", n.Text);
                cmd.Parameters.AddWithValue("@c", c.Text);
                cmd.Parameters.AddWithValue("@cd", cd.Text);
                cmd.Parameters.AddWithValue("@gc", gc.Text);
                cmd.Parameters.AddWithValue("@gu", gu.Text);
                cmd.Parameters.AddWithValue("@gcol", gcol.Text);
                cmd.Parameters.AddWithValue("@gd", gd.Text);
                cmd.Parameters.AddWithValue("@gm", gm.Text);
                cmd.Parameters.AddWithValue("@cm", cm.Text);
                cmd.Parameters.AddWithValue("@gnu", gnu.Text);
                cmd.Parameters.AddWithValue("@snn", snn.Text);
                cmd.Parameters.AddWithValue("@nnd", nnd.Text);
                cmd.Parameters.AddWithValue("@pn", pn.Text);
                conn.Open();
                cmd.ExecuteNonQuery();
                MessageBox.Show("Data Updated Successfully");
                conn.Close();
                DisplayData();
                ClearData();
            }

            else
            {
                MessageBox.Show("Error");
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Id = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
            fn.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            dn.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            gn.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            gdn.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            nn.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            s.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            mn.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            mdn.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
            mgn.Text = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();
            d.Text = dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString();
            n.Text = dataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString();
            c.Text = dataGridView1.Rows[e.RowIndex].Cells[12].Value.ToString();
            cd.Text = dataGridView1.Rows[e.RowIndex].Cells[13].Value.ToString();
            gc.Text = dataGridView1.Rows[e.RowIndex].Cells[14].Value.ToString();
            gu.Text = dataGridView1.Rows[e.RowIndex].Cells[15].Value.ToString();
            gcol.Text = dataGridView1.Rows[e.RowIndex].Cells[16].Value.ToString();
            gd.Text = dataGridView1.Rows[e.RowIndex].Cells[17].Value.ToString();
            gm.Text = dataGridView1.Rows[e.RowIndex].Cells[18].Value.ToString();
            cm.Text = dataGridView1.Rows[e.RowIndex].Cells[19].Value.ToString();
            gnu.Text = dataGridView1.Rows[e.RowIndex].Cells[20].Value.ToString();
            snn.Text = dataGridView1.Rows[e.RowIndex].Cells[21].Value.ToString();
            nnd.Text = dataGridView1.Rows[e.RowIndex].Cells[22].Value.ToString();
            pn.Text = dataGridView1.Rows[e.RowIndex].Cells[23].Value.ToString();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {

            DataGridViewRow row = dataGridView1.Rows[dataGridView1.CurrentRow.Index] as DataGridViewRow;
            if (row != null)
            {
                if (MessageBox.Show("Are you sure want to delete this record?", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    using (SqlConnection cn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\testdata.mdf;Integrated Security=True;"))
                    {
                        if (cn.State == ConnectionState.Closed)
                            cn.Open();
                        using (SqlCommand cmd = new SqlCommand("delete from babel where Id = @id", cn))
                        {
                            cmd.Parameters.AddWithValue("Id", row.Cells[0].Value);
                            cmd.ExecuteNonQuery();
                            DisplayData();
                            ClearData();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a row to delete");
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {

                    try


                    {
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(this.testdataDataSet1.babel.CopyToDataTable(), "babel");
                            workbook.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("Data successfully saved as an exel workbook!", "message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }


                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message, "message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }

                }

            }


        }

        private void search_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                using (SqlConnection cn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\testdata.mdf;Integrated Security=True;"))
                {
                    DataTable dt = new DataTable("Customer");
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("select * from babel where Fn like @Search", cn);
                    dataAdapter.SelectCommand.Parameters.AddWithValue("@Search", $"%{search.Text}%");
                    dataAdapter.Fill(dt);
                    dataGridView1.DataSource = dt;
                }
            }
        }

        DataTableCollection tableCollection;
        private void toolStripButton6_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filename.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                                cboSheet.Items.Add(table.TableName);

                        }
                    }
                }
            }
        }

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[cboSheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;

        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void search_TextChanged(object sender, EventArgs e)
        {
            if (searchcmbobox.Text == "الأسم الأول")
            {

                using (SqlConnection cn = new SqlConnection(url))
                {
                    DataTable dt = new DataTable("Customer");
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("select * from babel where Fn like @Search", cn);
                    dataAdapter.SelectCommand.Parameters.AddWithValue("@Search", $"%{search.Text}%");
                    dataAdapter.Fill(dt);
                    dataGridView1.DataSource = dt;
                }

            }


            else if (searchcmbobox.Text == "اللقب العلمي")
            {

                using (SqlConnection cn = new SqlConnection(url))
                {
                    DataTable dt = new DataTable("Customer");
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("select * from babel where Snn1 like @Search", cn);
                    dataAdapter.SelectCommand.Parameters.AddWithValue("@Search", $"%{search.Text}%");
                    dataAdapter.Fill(dt);
                    dataGridView1.DataSource = dt;
                }

            }

            else 
            {

                using (SqlConnection cn = new SqlConnection(url))
                {
                    DataTable dt = new DataTable("Customer");
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("select * from babel where C like @Search", cn);
                    dataAdapter.SelectCommand.Parameters.AddWithValue("@Search", $"%{search.Text}%");
                    dataAdapter.Fill(dt);
                    dataGridView1.DataSource = dt;
                }

            }
        }

        private void c_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {


            


        }





        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            numberToolStripMenuItem.Text = (dataGridView1.RowCount - 1).ToString();

            

        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            
        }

        private void toolStripDropDownButton1_Click(object sender, EventArgs e)
        {

        }

        private void عددالحاصلينعلىشهادةالبكالوريوسToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void عددالحاصلينعلىشهادةالبكالوريوسToolStripMenuItem_MouseHover(object sender, EventArgs e)
        {
            search.Text = "بكالوريوس";
            using (SqlConnection cn = new SqlConnection(url))
            {

                DataTable dt = new DataTable("Customer");
                SqlDataAdapter dataAdapter = new SqlDataAdapter("select * from babel where C like @Search", cn);
                dataAdapter.SelectCommand.Parameters.AddWithValue("@Search", $"%{search.Text}%");
                dataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }

            number2ToolStripMenuItem.Text = (dataGridView1.RowCount - 1).ToString();

        }

        private void عددالحاصلينعلىشهادةالبكالوريوسToolStripMenuItem_MouseLeave(object sender, EventArgs e)
        {
            DisplayData();
            search.Text = " ";
            
        }

        private void عددالحاصلينعلىشهادةالماجستيرToolStripMenuItem_MouseHover(object sender, EventArgs e)
        {
            search.Text = "ماجستير";
            using (SqlConnection cn = new SqlConnection(url))
            {

                DataTable dt = new DataTable("Customer");
                SqlDataAdapter dataAdapter = new SqlDataAdapter("select * from babel where C like @Search", cn);
                dataAdapter.SelectCommand.Parameters.AddWithValue("@Search", $"%{search.Text}%");
                dataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }

            number3ToolStripMenuItem.Text = (dataGridView1.RowCount - 1).ToString();

        }

        private void عددالحاصلينعلىشهادةالماجستيرToolStripMenuItem_MouseLeave(object sender, EventArgs e)
        {
            DisplayData();
            search.Text = " ";
        }

        private void عددالحاصلينعلىشهادةالدكتوراهToolStripMenuItem_MouseHover(object sender, EventArgs e)
        {
            search.Text = "دكتوراه";
            using (SqlConnection cn = new SqlConnection(url))
            {

                DataTable dt = new DataTable("Customer");
                SqlDataAdapter dataAdapter = new SqlDataAdapter("select * from babel where C like @Search", cn);
                dataAdapter.SelectCommand.Parameters.AddWithValue("@Search", $"%{search.Text}%");
                dataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }

            number4ToolStripMenuItem.Text = (dataGridView1.RowCount - 1).ToString();
        }

        private void عددالحاصلينعلىشهادةالدكتوراهToolStripMenuItem_MouseLeave(object sender, EventArgs e)
        {
            DisplayData();
            search.Text = " ";
        }
    }
}
