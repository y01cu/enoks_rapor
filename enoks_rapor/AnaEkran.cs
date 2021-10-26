using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace enoks_rapor
{
    public partial class AnaEkran : UserControl
    {
        SqlConnection connection = new SqlConnection("Server=DESKTOP-EE4KQN3\\UKO;Database=Deneme;Integrated Security=True");
        SqlDataAdapter dataAdapter;

        public AnaEkran()
        {
            InitializeComponent();
        }

        private void btnListele_Click(object sender, EventArgs e)
        {
            listing();
        }

        private void listing()
        {

            this.Cursor = Cursors.WaitCursor;
            string table = comboBox1.Text;
            string area = "";
            string area1 = "RAPOR SECINIZ";
            if (String.Equals(table, area) || string.Equals(table, area1))
            {
                MessageBox.Show("Lütfen alan seçiniz.", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = Cursors.Default;

            }
            else
            {
                string date1 = dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss.FFF");
                string date2 = dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss.FFF");
                connection.Open();
                dataAdapter = new SqlDataAdapter("select * from " + table + " WHERE TARIH BETWEEN '" + date1 + "' and  '" + date2 + "'", connection);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
                label10.Text = (dataGridView1.Rows.Count - 1).ToString();
                label11.Text = dataGridView1.Columns.Count.ToString();
                this.Cursor = Cursors.Default;
                dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
                dataGridView1.AutoSize = false;
                dataGridView1.ScrollBars = ScrollBars.Both;
                connection.Close();
            }
        }

        private void UserControl3_Load(object sender, EventArgs e)
        {
            cmbBox1();
        }

        private void cmbBox1()
        {
            try
            {
                connection.Open();
                DataTable dataTable = connection.GetSchema("Tables");
                int tableNumber = dataTable.Rows.Count;
                for (int i = 0; i < tableNumber; i++)
                {
                    comboBox1.Items.Add(dataTable.Rows[i]["TABLE_NAME"]);
                }
            }
            catch (Exception ex) { MessageBox.Show("Tablo görüntülerken hata meydana geldi." + ex, "HATA", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                connection.Close();
            }
        }

        private void dataProcess()
        {
            comboBox2.Items.Clear();
            string table = comboBox1.Text;
            connection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME=" + "'" + table + "'", connection);
            adapter.Fill(dataSet);
            for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
            {
                comboBox2.Items.Add(dataSet.Tables[0].Rows[i][0].ToString());
            }
            connection.Close();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd-MM-yyyy HH:mm:ss";
            dateTimePicker1.MinDate = new DateTime(2014, 07, 12);
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd-MM-yyyy HH:mm:ss";
            dateTimePicker2.MinDate = new DateTime(2014, 07, 12);
        }        

        private void deleting()
        {
            if (Giris.permission)
            {
                string table = comboBox1.Text;
                string area = "";
                string area1 = "DEGISKEN SECINIZ";
                if (String.Equals(table, area) || String.Equals(table, area1))
                {
                    MessageBox.Show("Lütfen alan seçiniz.", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DialogResult permission;
                    permission = MessageBox.Show("Seçmiş olduğunuz tarihler arasındaki kayıtları silmek istiyor musunuz?", "KAYIT SİLME", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (permission == DialogResult.Yes)
                    {
                        string date1 = dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss.FFF");
                        string date2 = dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss.FFF");
                        connection.Open();
                        SqlCommand deleteCommand = new SqlCommand("delete from " + table + " WHERE TARIH BETWEEN '" + date1 + "'and '" + date2 + "'", connection);
                        int etk = deleteCommand.ExecuteNonQuery();
                        if (etk > 0)
                        {
                            MessageBox.Show("Kayıtlar başarı ile silinmiştir.", "BİLGİLENDİRME", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Kayıt silme işleminde hata meydana geldi.", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        connection.Close();
                        listing();
                    }
                }
            }
            else
            {
                MessageBox.Show("Lütfen giriş yapınız.", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnKayıtSil_Click(object sender, EventArgs e)
        {
            deleting();
        }

        private void btnXlsKaydet_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "" || comboBox1.Text == "RAPOR SECINIZ")
            {
                MessageBox.Show("Öncelikle raporunuzu seçiniz.", "BİLGİLENDİRME", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                this.Cursor = Cursors.WaitCursor;
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }
                StartRow++;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            string gridtarih = dataGridView1[j, i].Value.ToString();
                            myRange.Value2 = dataGridView1[j, i].Value == null ? "" : gridtarih;
                        }
                        catch
                        { }
                    }
                }
                this.Cursor = Cursors.Default;
            }
        }

        private void datagrid2()
        {
            string table1 = comboBox1.Text;
            string colon = comboBox2.Text;
            string date1 = dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss.FFF");
            string date2 = dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss.FFF");
            SqlDataAdapter dataAdapter = new SqlDataAdapter("select TARIH," + colon + " from " + table1 + " WHERE TARIH BETWEEN'" + date1 + "' and '" + date2 + "'", connection);
            DataTable dataTable1 = new DataTable();
            dataAdapter.Fill(dataTable1);
            dataGridView2.DataSource = dataTable1;
        }

        private void sütunFark(DataGridViewCellEventArgs e)
        {
            double fark = Convert.ToDouble(dataGridView1[e.ColumnIndex, dataGridView1.Rows.Count - 2].Value) - Convert.ToDouble(dataGridView1[e.ColumnIndex, 0].Value);
            lblFark.Text = fark.ToString();
        }

        private void sütunOrtalama(DataGridViewCellEventArgs e)
        {
            double ortalama = Convert.ToDouble(sütunToplam(e) / (dataGridView1.Rows.Count - 1));
            lblSütun.Text = dataGridView1.Columns[e.ColumnIndex].HeaderText;
            lblOrtalama.Text = ortalama.ToString();
        }

        private double sütunToplam(DataGridViewCellEventArgs e)
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                toplam += Convert.ToDouble(dataGridView1[e.ColumnIndex, i].Value);
            }
            lblToplam.Text = toplam.ToString();
            return toplam;
        }

        private void btnYazdır_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bm = new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);
            dataGridView1.DrawToBitmap(bm, new System.Drawing.Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));
            e.Graphics.DrawImage(bm, 0, 0);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            lblSütun.Visible = true; chBoxFark.Visible = true; chBoxOrtalama.Visible = true; chBoxToplam.Visible = true; panelIslemler.Visible = true;
            panelSoldan1.Visible = true; panelSoldan2.Visible = true; panelSoldan3.Visible = true; panelSoldan4.Visible = true; panelSoldan5.Visible = true;
            try
            {
                dataGridView1.ClearSelection();
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1[e.ColumnIndex, i].Selected = true;
                }
                chBoxFark.Checked = false;
                chBoxToplam.Checked = false;
                chBoxOrtalama.Checked = false;
                lblToplam.Visible = false;
                lblFark.Visible = false;
                lblOrtalama.Visible = false;
                lblSütun.Text = dataGridView1.Columns[e.ColumnIndex].HeaderText;
                if (lblSütun.Text == "TARIH" || lblSütun.Text == null)
                {
                    lblSütun.Text = "GECERSIZ";
                }
                sütunToplam(e);
                sütunOrtalama(e);
                sütunFark(e);
            }

            catch (Exception)
            {
                MessageBox.Show("O sütunla işlem yapılmamaktadır.", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void chBoxToplam_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxToplam.Checked == false)
            {
                lblToplam.Visible = false;
            }
            else
            {
                lblToplam.Visible = true;
            }
        }

        private void chBoxFark_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxFark.Checked == false)
            {
                lblFark.Visible = false;
            }
            else
            {
                lblFark.Visible = true;
            }
        }

        private void chBoxOrtalama_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxOrtalama.Checked == false)
            {
                lblOrtalama.Visible = false;
            }
            else
            {
                lblOrtalama.Visible = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataProcess();
        }

        private void btnGrafikÇiz_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            String tablo = comboBox2.Text;
            string alan = "";
            if (String.Equals(tablo, alan) == true || String.Equals(tablo, "DEGISKEN SECINIZ") == true)
            {
                MessageBox.Show("Lütfen değişken seçiniz.", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = Cursors.Default;
            }
            else
            {
                try
                {
                    string tablo1 = comboBox1.Text;
                    string kolon = comboBox2.Text;
                    string tarih1 = dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss.FFF");
                    string tarih2 = dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss.FFF");
                    string query = ("select TARIH, " + kolon + " from " + tablo1 + " WHERE TARIH BETWEEN'" + tarih1 + "'and'" + tarih2 + "'");
                    SqlDataAdapter dr = new SqlDataAdapter(query, connection);
                    DataTable dt1 = new DataTable();
                    dr.Fill(dt1);
                    dataGridView2.DataSource = dt1;
                    foreach (var series in chart1.Series)
                    {
                        series.Points.Clear();
                    }
                    for (int i = 0; i < (dataGridView2.Rows.Count - 1); i++)
                    {
                        this.chart1.Series["DEGER"].Points.AddXY(dataGridView2.Rows[i].Cells[0].Value.ToString(), Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value.ToString()));
                    }
                    this.Cursor = Cursors.Default;
                    string maksimum_deg;
                    string minumum_deg;
                    string top_deg;
                    top_deg = dataGridView2.Rows[0].Cells[1].Value.ToString();
                    maksimum_deg = dataGridView2.Rows[0].Cells[1].Value.ToString();
                    minumum_deg = dataGridView2.Rows[0].Cells[1].Value.ToString();
                    float ort_deg;
                    float toplam_deg = float.Parse(top_deg);
                    float mini_deg = float.Parse(minumum_deg);
                    float maks_deg = float.Parse(maksimum_deg);
                    int sayi = dataGridView2.Rows.Count - 1;
                    for (int i = 1; i < (dataGridView2.Rows.Count - 1); i++)
                    {
                        string ara = dataGridView2.Rows[i].Cells[1].Value.ToString();
                        float ara1 = float.Parse(ara);
                        if (ara1 > maks_deg)
                        {
                            maks_deg = ara1;
                        }
                        if (ara1 < mini_deg)
                        {
                            mini_deg = ara1;
                        }
                        toplam_deg = toplam_deg + ara1;
                    }
                    ort_deg = toplam_deg / sayi;
                    txtMaksimum.Text = maks_deg.ToString();
                    txtMinimum.Text = mini_deg.ToString();
                    txtOrtalama.Text = ort_deg.ToString();
                }
                catch (Exception)
                {
                    MessageBox.Show("O sütunla işlem yapılmamaktadır.", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                this.Cursor = Cursors.Default;
            }
        }

        private void btnLogIn_Click(object sender, EventArgs e)
        {
            Giris gir = new Giris();
            gir.Show();
        }

        private void btnLogOut_Click(object sender, EventArgs e)
        {
            Giris.permission = false;
            MessageBox.Show("Çıkış yapıldı.", "BİLGİLENDİRME", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
