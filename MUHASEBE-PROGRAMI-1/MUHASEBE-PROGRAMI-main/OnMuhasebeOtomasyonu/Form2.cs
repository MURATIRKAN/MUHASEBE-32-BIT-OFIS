using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Windows.Forms;
using System.Xml;
namespace OnMuhasebeOtomasyonu
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }//

        public DataTable tablo = new DataTable();// tablo oluşturmak için datatable
        public DataTable tablo2 = new DataTable();
        OleDbConnection DATABASE_BAGLAN = new OleDbConnection("provider=microsoft.ace.oledb.16.0;data source=" + Application.StartupPath + "\\datam.accdb");

        //bağlantı kodumuzu hazırladık
        private void button1_Click(object sender, EventArgs e)
        {
            panel2.Hide();//anasayfadaki panel geçişlerini ayarladık
            panel1.Show();
            panel3.Hide();

        }


        private void button2_Click(object sender, EventArgs e)
        {
            Form3 FRM3 = new Form3();
            FRM3.Show();//form3 geçtik

        }
        public void listele()
        {
            try
            {
                tablo.Clear();//tabloyu temizledik
                DATABASE_BAGLAN.Open();//bağlantı açtık
                OleDbDataAdapter adtr = new OleDbDataAdapter("select * From stokbil", DATABASE_BAGLAN);
                adtr.Fill(tablo);//stokları listeleyip datatable doldurduk
                dataGridView1.DataSource = tablo;//datatabledan datagridwiewe aktardık
                adtr.Dispose();
                DATABASE_BAGLAN.Close();//bağlantıyı kapadık
                try
                {
                    dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //datagridview1'deki tüm satırı seç              
                    dataGridView1.Columns[1].HeaderText = "STOK ADI";
                    //sütunlardaki textleri değiştirme
                    dataGridView1.Columns[2].HeaderText = "STOK MODELİ";
                    dataGridView1.Columns[3].HeaderText = "STOK SERİNO";
                    dataGridView1.Columns[4].HeaderText = "STOK ADEDİ";
                    dataGridView1.Columns[5].HeaderText = "STOK TARİH";
                    dataGridView1.Columns[6].HeaderText = "KAYIT YAPAN";
                    dataGridView1.Columns[7].HeaderText = "DOSYA ADI";
                    dataGridView1.Columns[8].HeaderText = "ÖLÇÜ BİRİMİ";
                    dataGridView1.Columns[9].HeaderText = "ALIŞ FİYAT";
                    dataGridView1.Columns[10].HeaderText = "SATIŞ FİYAT";//tablo başlık isimlerini düzenledik
                    dataGridView1.Columns[0].Visible = false;//başlangiç kolununu gizledik
                    dataGridView1.Columns[1].Width = 120;
                    //genişlik
                    dataGridView1.Columns[2].Width = 120;
                    dataGridView1.Columns[3].Width = 120;
                    dataGridView1.Columns[4].Width = 80;
                    dataGridView1.Columns[5].Width = 100;
                    dataGridView1.Columns[6].Width = 120;//datagridwiewde bazı alanların genişliği ayarladık
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

        void dovizdoldur()
        {
            XmlDocument xmlVerisi = new XmlDocument();
            xmlVerisi.Load("http://www.tcmb.gov.tr/kurlar/today.xml");

            decimal dolar = Convert.ToDecimal(xmlVerisi.SelectSingleNode(string.Format("Tarih_Date/Currency[@Kod='{0}']/ForexSelling", "USD")).InnerText.Replace('.', ','));

            decimal Euro = Convert.ToDecimal(xmlVerisi.SelectSingleNode(string.Format("Tarih_Date/Currency[@Kod='{0}']/ForexSelling", "EUR")).InnerText.Replace('.', ','));

            label1.Text = dolar.ToString();
            label2.Text = Euro.ToString();
        }
        private void Form2_Load(object sender, EventArgs e)
        {

            dovizdoldur();


            listele();

            OleDbCommand cmd = new OleDbCommand();
            DATABASE_BAGLAN.Open();
            cmd.Connection = DATABASE_BAGLAN;
            cmd.CommandText = "SELECT * FROM hareket";
            OleDbDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                listBox1.Items.Add(dr["hareket"].ToString() + dr["tarih"].ToString() + dr["kullanici".ToString()]);


            }


            DATABASE_BAGLAN.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            panel2.Show();
            panel3.Hide();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form5 frm5 = new Form5();
            frm5.Show();


        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form6 frm6 = new Form6();
            frm6.Show();


        }


        private void button3_Click(object sender, EventArgs e)
        {
            Form4 frm4 = new Form4();
            frm4.Show();


        }

        private void çıkışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Bizi Tercih Ettiğiniz İçin Teşekkür Ederiz!");
            Application.Exit();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Process Process = new Process();//hesap makinesini çalişdıracak kodları girdik
            ProcessStartInfo ProcessInfo;
            ProcessInfo = new ProcessStartInfo("cmd.exe", "/C " + "calc");
            ProcessInfo.CreateNoWindow = true;
            ProcessInfo.UseShellExecute = false;

            Process = Process.Start(ProcessInfo);
            Process.WaitForExit();
            Process.Close();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void stokEklemeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.Show();
        }

        private void hareketlerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel2.Show();
            panel3.Hide();
        }

        private void cariİşlemlerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form5 frm5 = new Form5();
            frm5.Show();

        }

        private void ajandaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form6 frm6 = new Form6();
            frm6.Show();
        }

     
        private void button9_Click(object sender, EventArgs e)
        {//stok işlemleri yaptımız forma geçer
            Form3 frm3 = new Form3();
            frm3.Show();
        }

        private void kesilenFaturalarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form5 frm5 = new Form5();
            frm5.Show();
        }

      
        private void button11_Click(object sender, EventArgs e)
        {
            kullanici musteri = new kullanici();
            musteri.Show();
            this.Hide();
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            kasa_Hesap_takip kasa = new kasa_Hesap_takip();
            kasa.Show();
            this.Hide();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            cek_Senet cek = new cek_Senet();
            cek.Show();
            this.Hide();
        }
    }
}
