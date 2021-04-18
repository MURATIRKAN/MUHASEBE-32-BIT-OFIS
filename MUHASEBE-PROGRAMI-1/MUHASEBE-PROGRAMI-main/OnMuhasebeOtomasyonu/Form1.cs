//****************************************************************
//     MICROSOFT C# ÖN MUHASEBE PROGRAMI v2.03 OFİS 2016(x86)
// (WINDOWS İŞLETİM SİSTEMİ 10 VEYA 8) MURAT IRKAN 2021-02
//****************************************************************
using System;
using System.Data.OleDb;
using System.Net.NetworkInformation;
using System.Windows.Forms;
namespace OnMuhasebeOtomasyonu
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //  DATABASE_BAGLAN = new OleDbConnection("provider=microsoft.ace.oledb.16.0;data source=" + Application.StartupPath + "\\DATABASE\\datam.accdb");
        OleDbConnection DATABASE_BAGLAN = new OleDbConnection("provider=microsoft.ace.oledb.16.0;data source=" + Application.StartupPath + "\\datam.accdb");

        private void Form1_Load(object sender, EventArgs e)
        {

            string program_adı = "Microsoft Office";
            bool prg_kontrol = OFIS_KONTROL(program_adı);
            if (prg_kontrol)
                MessageBox.Show("MICROSOFT OFİS SİSTEMDE YÜKLÜ");
            else
                MessageBox.Show("MICROSOFT OFIS 32-BIT YÜKLÜ DEGİL.RAPORLAR ÇALIŞAMAZ.!");

            Microsoft.Win32.RegistryKey key;
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Names");
            key.SetValue("ÖNMUHASEBE", DateTime.Now);
            key.Close();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void button1_Click(object sender, EventArgs e)
        {
     
            if ((textBox1.Text == "") & (textBox2.Text == ""))
            {
                MessageBox.Show("GİRİŞ BOŞ HATASI.TEKRAR DENEYİN.!");
            }
            if (NetworkInterface.GetIsNetworkAvailable() == true)
            {
                MessageBox.Show("İNTERNET BAGLANTISI ÇALIŞIYOR.!");
            }
            else
            {
                MessageBox.Show("İNTERNET BAĞLANTI SORUNU.!", "LÜTFEN KONTROL EDİN.?",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            DATABASE_BAGLAN.Open();// OFİS DATABASE BAGLANTISINI AC
            OleDbCommand DATABASE_KOMUT = new OleDbCommand("Select * from kullanici where k_adi='" + textBox1.Text + "'and sifre='" + textBox2.Text + "'", DATABASE_BAGLAN);
            OleDbDataReader DATABASE_OKUYUCU = DATABASE_KOMUT.ExecuteReader();
            if (DATABASE_OKUYUCU.Read()) // EYER DATABASE OKUNDUYSA
            {

                Form2 frm2 = new Form2();// ANA  FORMU TANIMLA
                frm2.Show();// ANA MUHASEBE EKRANI GOSTER.!
                this.Hide();// Bİ ONCEKI FORMU KAPAT.!
                MessageBox.Show("C# ÖN MUHASEBE HOŞGELDİNİZ.!");
            }
            else
            {
                MessageBox.Show("ÜYE KAYIT BULUNAMADI.!", "LÜTFEN KONTROL EDİN.?",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            DATABASE_BAGLAN.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            iletisim ilet = new iletisim();
            ilet.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("iletisim bilgilerim: muratir1975@windowslive.com");
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            sifremiunuttum sifre = new sifremiunuttum();
            sifre.Show();
            this.Hide();
        }
        public static bool OFIS_KONTROL(string PROGRAM_ADI)
        {
            bool PRG_DURUMU = false;
            string registry_ANAHTAR = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(registry_ANAHTAR))
            {
                foreach (string subkey_name in key.GetSubKeyNames())
                {
                    using (Microsoft.Win32.RegistryKey subkey = key.OpenSubKey(subkey_name))
                    {
                        if (!string.IsNullOrEmpty(Convert.ToString(subkey.GetValue("DisplayName"))))
                        {
                            if (Convert.ToString(subkey.GetValue("DisplayName")).Contains(PROGRAM_ADI))
                                PRG_DURUMU = true;
                        }
                    }
                }
            }
            return PRG_DURUMU;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            uyeOl uye = new uyeOl();
            uye.Show();
            this.Hide();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("ÖN MUHASEBE KULLANDIĞINIZ İÇİN TŞK EDERİZ.");
            Application.Exit();

        }
    }
}
