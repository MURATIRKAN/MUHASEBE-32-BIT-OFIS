using System;
using System.Data.OleDb;
using System.Windows.Forms;
namespace OnMuhasebeOtomasyonu
{
    public partial class sifremiunuttum : Form
    {
        public sifremiunuttum()
        {
            InitializeComponent();
        }
        OleDbConnection baglan = new OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\datam.accdb");

        private void button1_Click(object sender, EventArgs e)
        {
            doldur();
        }

        public void doldur()
        {

            baglan.Open();
            OleDbCommand kmt = new OleDbCommand("Select * from kullanici where k_adi='" + textBox1.Text + "'and gizlisoru='" + comboBox1.Text + "'and yaniti='" + textBox2.Text + "'", baglan);
            OleDbDataReader dr = kmt.ExecuteReader();
            if (dr.Read())
            {
                textBox3.Text = dr["sifre"].ToString();



            }
            else
            {
                MessageBox.Show("Bilgiler uyuşmuyor");
            }
            baglan.Close();
        }
    }
}
