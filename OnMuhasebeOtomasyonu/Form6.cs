using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
namespace OnMuhasebeOtomasyonu
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }
        public DataTable DATA_TABLOSU = new DataTable();
        public OleDbConnection DATABASE_BAGLAN = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=datam.accdb");
        public OleDbDataAdapter DATA_ADAPTORU = new OleDbDataAdapter();
        public OleDbCommand DATABASE_KOMUTU = new OleDbCommand();

        private void Form6_Load(object sender, EventArgs e)
        {
            textBox3.Text = DateTime.Now.ToShortDateString();
            listele();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DATABASE_BAGLAN.Open();
            DATABASE_KOMUTU.Connection = DATABASE_BAGLAN;
            DATABASE_KOMUTU.CommandText = "INSERT INTO notlar(adi,not_,ktarih,gtarih) VALUES ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + maskedTextBox1.Text + "')";
            DATABASE_KOMUTU.ExecuteNonQuery();


            DATABASE_BAGLAN.Close();
            listele();
            MessageBox.Show("Ajandanıza Notunuz Başarıyla Kaydedildi");
        }
        void listele()
        {
            DATA_TABLOSU.Clear();
            DATABASE_BAGLAN.Open();
            OleDbDataAdapter DATA_ADAPTORU = new OleDbDataAdapter("select * From notlar", DATABASE_BAGLAN);
            DATA_ADAPTORU.Fill(DATA_TABLOSU);
            dataGridView1.DataSource = DATA_TABLOSU;
            DATA_ADAPTORU.Dispose();
            DATABASE_BAGLAN.Close();
            dataGridView1.Columns[1].HeaderText = "NOT ADI";
            //sütunlardaki textleri değiştirme
            dataGridView1.Columns[2].HeaderText = "NOT İÇERİĞİ";
            dataGridView1.Columns[3].HeaderText = "NOT KAYIT TARİHİ";
            dataGridView1.Columns[4].HeaderText = "NOT GÖSTERİM TARİHİ";
            dataGridView1.Columns[2].Width = 240;
            dataGridView1.Columns[4].Width = 140;
            dataGridView1.Columns[3].Width = 140;
            dataGridView1.Columns[0].Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
