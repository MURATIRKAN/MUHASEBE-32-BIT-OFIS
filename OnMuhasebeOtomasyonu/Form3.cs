using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace OnMuhasebeOtomasyonu
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        string DosyaYolu, DosyaAdi = "";
        public DataTable tablo = new DataTable();
        public OleDbConnection bag = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=datam.accdb");
        public OleDbDataAdapter adtr = new OleDbDataAdapter();
        public OleDbCommand kmt = new OleDbCommand();
        int id;
        int SAY = 0;
        private void Form3_Load(object sender, EventArgs e)
        {
            listele();
        }

        private void btnResimEkle_Click(object sender, EventArgs e)
        {

            if (DosyaAc.ShowDialog() == DialogResult.OK)
            {
                foreach (string i in DosyaAc.FileName.Split('\\'))
                {
                    if (i.Contains(".jpg")) { DosyaAdi = i; }
                    else if (i.Contains(".png")) { DosyaAdi = i; }
                    else { DosyaYolu += i + "\\"; }
                }
                pictureBox1.ImageLocation = DosyaAc.FileName;
            }
            else
            {
                MessageBox.Show("Dosya Girmediniz!");
            }
        }

        private void btnResimSil_Click(object sender, EventArgs e)
        {
            pictureBox1.ImageLocation = "";
            DosyaAdi = "";
        }

        private void btnStokEkle_Click(object sender, EventArgs e)
        {
            // KAYIT SAYISI
            SAY = tablo.Rows.Count;
            if (SAY > 5)
            {
                MessageBox.Show("SINIRLI KAYIT.!LÜFTEN TAM SÜRÜMÜ İSTEYİNİZ.!", "SINIRLI KULLANIM BİLGİSİ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            try
            {
                if (textBox1.Text.Trim() == "") errorProvider1.SetError(textBox1, "Boş geçilmez");
                else errorProvider1.SetError(textBox1, "");
                if (textBox2.Text.Trim() == "") errorProvider1.SetError(textBox2, "Boş geçilmez");
                else errorProvider1.SetError(textBox2, "");
                if (textBox3.Text.Trim() == "") errorProvider1.SetError(textBox3, "Boş geçilmez");
                else errorProvider1.SetError(textBox3, "");
                if (textBox4.Text.Trim() == "") errorProvider1.SetError(textBox4, "Boş geçilmez");
                else errorProvider1.SetError(textBox4, "");
                if (textBox5.Text.Trim() == "") errorProvider1.SetError(textBox5, "Boş geçilmez");
                else errorProvider1.SetError(textBox5, "");

                if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "" && textBox3.Text.Trim() != "" && textBox4.Text.Trim() != "" && textBox5.Text.Trim() != "")
                {
                    bag.Open();
                    kmt.Connection = bag;
                    kmt.CommandText = "INSERT INTO stokbil(stokAdi,stokModeli,stokSeriNo,stokAdedi,stokTarih,kayitYapan,dosyaAdi,olcu_birimi,afiyat,sfiyat) VALUES ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + dateTimePicker1.Text + "','" + textBox5.Text + "','" + DosyaAdi + "','" + comboBox1.Text + "','" + textBox7.Text + "','" + textBox8.Text + "') ";
                    kmt.ExecuteNonQuery();
                    kmt.Dispose();
                    bag.Close();
                    for (int i = 0; i < this.Controls.Count; i++)
                    {
                        if (this.Controls[i] is TextBox) this.Controls[i].Text = "";
                    }
                    listele();
                    if (DosyaAdi != "") File.WriteAllBytes(DosyaAdi, File.ReadAllBytes(DosyaAc.FileName));
                    //  MessageBox.Show("Kayıt İşlemi Tamamlandı ! ", "İşlem Sonucu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                bag.Open();
                kmt.Connection = bag;
                kmt.CommandText = "INSERT INTO hareket(hareket,tarih,kullanici) VALUES ('" + "STOK Ekleme İşlemi Yapılmıştır...Tarihi= " + "','" + DateTime.Now.ToLongDateString() + "','" + " İşlemi Gerçekleştiren= " + textBox5.Text + "') ";
                kmt.ExecuteNonQuery();


                bag.Close();

            }

            catch (Exception hata)
            {
                MessageBox.Show("Kayıtlı Seri No !" + hata);
                bag.Close();
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
                    kmt.CommandText = "DELETE from stokbil WHERE stokSeriNo='" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "' ";
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
            bag.Open();
            kmt.Connection = bag;
            kmt.CommandText = "INSERT INTO hareket(hareket,tarih,kullanici) VALUES ('" + "STOK Silme İşlemi Yapılmıştır...Tarihi= " + "','" + DateTime.Now.ToLongDateString() + "','" + " İşlemi Gerçekleştiren= " + textBox5.Text + "') ";
            kmt.ExecuteNonQuery();


            bag.Close();
        }

        private void btnStokGuncelle_Click(object sender, EventArgs e)
        {


            if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "" && textBox3.Text.Trim() != "" && textBox4.Text.Trim() != "" && textBox5.Text.Trim() != "")
            {


                string sorgu = "UPDATE stokbil SET stokAdi='" + textBox1.Text + "',stokModeli='" + textBox2.Text + "',stokSeriNo='" + textBox3.Text + "',stokAdedi='" + textBox4.Text + "',stokTarih='" + dateTimePicker1.Text + "',kayitYapan='" + textBox5.Text + "',dosyaAdi='" + DosyaAdi + "',olcu_birimi='" + comboBox1.Text + "',afiyat='" + textBox7.Text + "',sfiyat='" + textBox8.Text + "' WHERE id=" + id;
                OleDbCommand kmt = new OleDbCommand(sorgu, bag);
                bag.Open();
                kmt.ExecuteNonQuery();
                kmt.Dispose();
                bag.Close();
                listele();
                if (DosyaAdi != "") File.WriteAllBytes(DosyaAdi, File.ReadAllBytes(DosyaAc.FileName));
                MessageBox.Show("Güncelleme İşlemi Tamamlandı !", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Boş Alan Bırakmayınız !");
            }
            bag.Open();
            kmt.Connection = bag;
            kmt.CommandText = "INSERT INTO hareket(hareket,tarih,kullanici) VALUES ('" + "STOK Güncelleme İşlemi Yapılmıştır... Tarihi=  " + "','" + DateTime.Now.ToLongDateString() + "','" + " İşlemi Gerçekleştiren= " + textBox5.Text + "') ";
            kmt.ExecuteNonQuery();


            bag.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form4 frm3 = new Form4();
            frm3.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 FRM2 = new Form2();
            FRM2.Show();
            this.Hide();

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            //dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox7.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            textBox8.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            try
            {
                kmt = new OleDbCommand("select * from stokbil where stokSeriNo='" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "'", bag);
                bag.Open();
                OleDbDataReader oku = kmt.ExecuteReader();
                oku.Read();
                if (oku.HasRows)
                {
                    pictureBox1.ImageLocation = oku[7].ToString();
                    id = Convert.ToInt32(oku[0].ToString());
                }
                bag.Close();
            }
            catch
            {
                bag.Close();
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

            for (i = 0; i <= dataGridView1.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGridView1[j, i];
                    xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }

            xlWorkBook.SaveAs("STOKRAPORU.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            OBJE_OLUSTUR(xlWorkSheet);
            OBJE_OLUSTUR(xlWorkBook);
            OBJE_OLUSTUR(xlApp);
            MessageBox.Show(" OTOMATİK Excel Raporu oluştu");

        }
        private void OBJE_OLUSTUR(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("DOSYA OLUŞURKEN BİR HATA.! " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application EXCEL_UYGULAMASI;
            Excel.Workbook EXCEL_KİTAP;
            Excel.Worksheet EXCEL_SAYFASI;
            object HATA_OKUMASI = System.Reflection.Missing.Value;

            EXCEL_UYGULAMASI = new Excel.Application();
            EXCEL_KİTAP = EXCEL_UYGULAMASI.Workbooks.Add(HATA_OKUMASI);
            EXCEL_SAYFASI = (Excel.Worksheet)EXCEL_KİTAP.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;

            for (i = 0; i <= dataGridView1.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    DataGridViewCell hücre = dataGridView1[j, i];
                    EXCEL_SAYFASI.Cells[i + 1, j + 1] = hücre.Value;
                }
            }
            try
            {

                EXCEL_KİTAP.SaveAs("STOKRAPORU.xls", Excel.XlFileFormat.xlWorkbookNormal, HATA_OKUMASI, HATA_OKUMASI, HATA_OKUMASI, HATA_OKUMASI, Excel.XlSaveAsAccessMode.xlExclusive, HATA_OKUMASI, HATA_OKUMASI, HATA_OKUMASI, HATA_OKUMASI, HATA_OKUMASI);
                EXCEL_KİTAP.Close(true, HATA_OKUMASI, HATA_OKUMASI);
                EXCEL_UYGULAMASI.Quit();

                OBJE_OLUSTUR(EXCEL_SAYFASI);
                OBJE_OLUSTUR(EXCEL_KİTAP);
                OBJE_OLUSTUR(EXCEL_UYGULAMASI);

                MessageBox.Show("Excel Raporunuz Oluştu.! DOCUMENTS(DÖKÜNAMLARIM KLASÖRÜNDE STOKRAPORU.XLS.!");
                string DOSYA_EXCEL;
                DOSYA_EXCEL = @"STOKRAPORU.xls";
                Excel.Application xlapp;
                Excel.Workbook xlworkbook;
                xlapp = new Excel.Application();
                xlworkbook = xlapp.Workbooks.Open(DOSYA_EXCEL, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlapp.Visible = true;
            }

            catch (Exception HATA)
            {
                MessageBox.Show("KONTROL EDİN.! EXCEL DOSYA KAYIT HATASI.!", HATA.Message);
            }
        }
        public void listele()
        {
            try
            {
                tablo.Clear();
                bag.Open();
                OleDbDataAdapter adtr = new OleDbDataAdapter("select stokAdi,stokModeli,stokSeriNo,stokAdedi,stokTarih,kayitYapan,olcu_birimi,afiyat,sfiyat From stokbil", bag);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;
                adtr.Dispose();
                bag.Close();
                try
                {
                    dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //datagridview1'deki tüm satırı seç              
                    dataGridView1.Columns[0].HeaderText = "STOK ADI";
                    //sütunlardaki textleri değiştirme
                    dataGridView1.Columns[1].HeaderText = "STOK MODELİ";
                    dataGridView1.Columns[2].HeaderText = "STOK SERİNO";
                    dataGridView1.Columns[3].HeaderText = "STOK ADEDİ";
                    dataGridView1.Columns[4].HeaderText = "STOK TARİH";
                    dataGridView1.Columns[5].HeaderText = "KAYIT YAPAN";
                    dataGridView1.Columns[0].Width = 120;
                    //genişlik
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
    }
}
