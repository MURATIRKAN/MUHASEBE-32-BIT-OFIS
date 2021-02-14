using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OnMuhasebeOtomasyonu
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        public DataTable tablo = new DataTable();
        public DataTable tablo2 = new DataTable();
        public OleDbConnection bag = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=datam.accdb");
        public OleDbDataAdapter adtr = new OleDbDataAdapter();
        public OleDbCommand kmt = new OleDbCommand();
#pragma warning disable CS0169 // The field 'Form5.id' is never used
        int id;
        int SAY;
#pragma warning restore CS0169 // The field 'Form5.id' is never used

        private void Form5_Load(object sender, EventArgs e)
        {

            listele();
            doldur();


        }

        public void doldur()
        {
            bag.Open();
            kmt = new OleDbCommand("Select * from stokbil", bag);
            OleDbDataReader dr = kmt.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr["stokSeriNo"].ToString());



            }
            bag.Close();
        }
        public void listele()
        {
            try
            {
                tablo.Clear();
                bag.Open();
                OleDbDataAdapter adtr = new OleDbDataAdapter("select * From cari", bag);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;
                adtr.Dispose();
                bag.Close();
                try
                {
                    dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //datagridview1'deki tüm satırı seç              
                    dataGridView1.Columns[1].HeaderText = "STOK SERİNO";
                    //sütunlardaki textleri değiştirme
                    dataGridView1.Columns[2].HeaderText = "MÜŞTERİ ADI";
                    dataGridView1.Columns[3].HeaderText = "ÖDEME TİPİ";
                    dataGridView1.Columns[4].HeaderText = "ALIM MİKTARI";
                    dataGridView1.Columns[5].HeaderText = "BİRİM SATIŞ FİYATI";
                    dataGridView1.Columns[6].HeaderText = "AÇIKLAMA";
                    dataGridView1.Columns[7].HeaderText = "TARİH";
                    dataGridView1.Columns[8].Width = 120;
                    //genişlik
                    dataGridView1.Columns[0].Width = 0;
                    dataGridView1.Columns[1].Width = 120;
                    dataGridView1.Columns[2].Width = 120;
                    dataGridView1.Columns[3].Width = 80;
                    dataGridView1.Columns[4].Width = 100;
                    dataGridView1.Columns[5].Width = 120;
                }
                catch
                {
                    ;
                }
            }
            catch (Exception hata)
            {

                MessageBox.Show(hata.ToString());
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            bag.Open();
            kmt = new OleDbCommand("Select * from stokbil where stokSeriNo='" + comboBox1.Text + "'", bag);
            OleDbDataReader dr = kmt.ExecuteReader();
            while (dr.Read())
            {
                textBox6.Text = dr["sfiyat"].ToString();
            }
            bag.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            this.Hide();
        }

        private void btnStokEkle_Click(object sender, EventArgs e)
        {
            SAY = tablo.Rows.Count;
            if (SAY > 5)
            {
                MessageBox.Show("SINIRLI KAYIT.!LÜFTEN TAM SÜRÜMÜ İSTEYİNİZ.!", "SINIRLI KULLANIM BİLGİSİ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (textBox2.Text.Trim() != "" && textBox4.Text.Trim() != "" && textBox5.Text.Trim() != "")
            {
                bag.Open();
                kmt.Connection = bag;
                kmt.CommandText = "INSERT INTO cari(serino,musteri_adi,odeme_tipi,alim_miktari,tutar,aciklama,tarih) VALUES ('" + comboBox1.Text + "','" + textBox2.Text + "','" + comboBox2.Text + "','" + textBox4.Text + "','" + textBox6.Text + "','" + textBox5.Text + "','" + dateTimePicker1.Text + "') ";
                kmt.ExecuteNonQuery();
                kmt.Dispose();
                bag.Close();
                listele();

                MessageBox.Show("Kayıt İşlemi Tamamlandı ! ", "İşlem Sonucu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void btnStokSil_Click(object sender, EventArgs e)
        {

            try
            {
                DialogResult cevap;
                cevap = MessageBox.Show("Kaydı silmek istediğinizden eminmisiniz", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes && dataGridView1.CurrentRow.Cells[0].Value.ToString().Trim() != "")
                {
                    bag.Open();
                    kmt.Connection = bag;
                    kmt.CommandText = "DELETE from cari WHERE serino='" + dataGridView1.CurrentRow.Cells[1].Value.ToString() + "' ";
                    kmt.ExecuteNonQuery();
                    kmt.Dispose();
                    bag.Close();
                    listele();
                }
            }
            catch
            {
                ;
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                comboBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                comboBox2.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                textBox6.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                textBox5.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            }
            catch
            {

                ;
            }


        }

        private void btnStokGuncelle_Click(object sender, EventArgs e)
        {

            if (comboBox1.Text.Trim() != "" && textBox5.Text.Trim() != "")
            {


                string sorgu = "UPDATE cari SET serino='" + comboBox1.Text + "',musteri_adi='" + textBox2.Text + "',odeme_tipi='" + comboBox2.Text + "',alim_miktari='" + textBox4.Text + "',tutar='" + textBox6.Text + "',aciklama='" + textBox5.Text + "',tarih='" + dateTimePicker1.Text + "' WHERE id=" + dataGridView1.CurrentRow.Cells[0].Value.ToString();
                OleDbCommand kmt = new OleDbCommand(sorgu, bag);
                bag.Open();
                kmt.ExecuteNonQuery();
                kmt.Dispose();
                bag.Close();
                listele();
                MessageBox.Show("Güncelleme İşlemi Tamamlandı !", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Boş Alan Bırakmayınız !");
            }
        }

        private void btnStokModelAra_Click(object sender, EventArgs e)
        {
            bag.Open();
            OleDbDataAdapter adtr = new OleDbDataAdapter("select * From cari", bag);
            if (radioButton1.Checked == true)
            {
                if (textBox1.Text.Trim() == "")
                {
                    tablo.Clear();
                    kmt.Connection = bag;
                    kmt.CommandText = "Select * from cari";
                    adtr.SelectCommand = kmt;
                    adtr.Fill(tablo);
                }
                if (Convert.ToBoolean(bag.State) == false)
                {
                    bag.Open();
                }
                if (textBox1.Text.Trim() != "")
                {
                    adtr.SelectCommand.CommandText = " Select * From cari" +
                         " where(serino='" + textBox1.Text + "' )";
                    tablo.Clear();
                    adtr.Fill(tablo);
                    bag.Close();
                }


            }
            else if (radioButton2.Checked == true)
            {
                if (textBox1.Text.Trim() == "")
                {
                    tablo.Clear();
                    kmt.Connection = bag;
                    kmt.CommandText = "Select * from cari";
                    adtr.SelectCommand = kmt;
                    adtr.Fill(tablo);
                }
                if (Convert.ToBoolean(bag.State) == false)
                {
                    bag.Open();
                }
                if (textBox1.Text.Trim() != "")
                {
                    adtr.SelectCommand.CommandText = " Select * From cari" +
                         " where(musteri_adi='" + textBox1.Text + "' )";
                    tablo.Clear();
                    adtr.Fill(tablo);
                    bag.Close();
                }
            }
            else
            {
                MessageBox.Show("Lütfen bir arama türü seçiniz...");
            }
        }
        private void OBJE_OLUSTUR(object obje)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obje);
                obje = null;
            }
            catch (Exception hata)
            {
                obje = null;
                MessageBox.Show("RAPORLAMA SIRASINDA BİR HATA OLUŞTU.! " + hata.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void KONTROL(object sender, EventArgs e)
        {
            FileInfo dosya_kontrol = new FileInfo("C:\\Documents\\STOKRAPORU.xls");
            if (dosya_kontrol.Exists)
            {
                MessageBox.Show("RAPOR MYDOCUMENTS VEYA DOKÜMANLARIM KLASÖRÜNDE OLUŞTU.!");
            }
            else
            {
                MessageBox.Show("Sistemde Rapor bulunamadı.! KONTROL EDİN.?");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;
            try
            {
                for (i = 0; i <= dataGridView1.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGridView1[j, i];
                        xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                    }
                }
                // DOSYA KONTROL .!
                xlWorkBook.SaveAs("CARİ.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                OBJE_OLUSTUR(xlWorkSheet);
                OBJE_OLUSTUR(xlWorkBook);
                OBJE_OLUSTUR(xlApp);

                string fileExcel;
                fileExcel = @"CARİ.xls";
                Excel.Application xlapp;
                Excel.Workbook xlworkbook;
                xlapp = new Excel.Application();

                xlworkbook = xlapp.Workbooks.Open(fileExcel, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                xlapp.Visible = true;
            }
            catch (Exception HATA)
            { MessageBox.Show("RAPORLAMA OLUŞURKEN BİR HATA OLUŞTU.!", HATA.Message); }
        }




    }
}



