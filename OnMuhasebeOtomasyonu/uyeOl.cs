using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
namespace OnMuhasebeOtomasyonu
{
    public partial class uyeOl : Form
    {
        public uyeOl()
        {
            InitializeComponent();
        }
        OleDbConnection baglan = new OleDbConnection("provider=microsoft.ace.oledb.16.0;data source=" + Application.StartupPath + "\\datam.accdb");

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox6.Text == textBox5.Text)
            {
                try
                {
                    if (baglan.State == ConnectionState.Open)
                    {

                        baglan.Close();
                    }
                    baglan.Open();

                    OleDbCommand kmt;

                    kmt = new OleDbCommand("INSERT INTO kullanici(tc,adsoyad,tel,k_adi,sifre,gizlisoru,yaniti) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + comboBox1.Text + "','" + textBox7.Text + "')", baglan);
                    kmt.ExecuteNonQuery();
                    baglan.Close();
                    MessageBox.Show("Kayıt Başarılı");
                    Form1 frm2 = new Form1();
                    frm2.Show();
                }
                catch (Exception hata)
                {

                    MessageBox.Show(hata.Message);
                }
            }
            else
            {
                MessageBox.Show("Şifreler Örtüşmüyor.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
