using System;
using System.Data;
using System.Windows.Forms;
#pragma warning disable CS0105 // The using directive for 'System.Data' appeared previously in this namespace
#pragma warning restore CS0105 // The using directive for 'System.Data' appeared previously in this namespace
using System.Data.OleDb;
namespace OnMuhasebeOtomasyonu
{
    public partial class cek_Senet : Form
    {
        public cek_Senet()
        {
            InitializeComponent();
        }
        OleDbConnection DATABASE_BAGLAN = new OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\datam.accdb");
        DataTable DATA_TABLOSU = new DataTable();
        int SAY;
        public void listele()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            comboBox2.Text = "";


            DATA_TABLOSU.Clear();
            DATABASE_BAGLAN.Open();
            OleDbDataAdapter adtr = new OleDbDataAdapter
          ("select * from ceksenet", DATABASE_BAGLAN);
            adtr.Fill(DATA_TABLOSU);
            dataGridView1.DataSource = DATA_TABLOSU;
            dataGridView1.Columns[0].Visible = false;
            DATABASE_BAGLAN.Close();
        }

        private void cek_Senet_Load(object sender, EventArgs e)
        {
            listele();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SAY = DATA_TABLOSU.Rows.Count;
            if (SAY > 5)
            {
                MessageBox.Show("SINIRLI KAYIT.!LÜFTEN TAM SÜRÜMÜ İSTEYİNİZ.!", "SINIRLI KULLANIM BİLGİSİ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            try
            {
                if (DATABASE_BAGLAN.State == ConnectionState.Open)
                {
                    DATABASE_BAGLAN.Close();
                }
                DATABASE_BAGLAN.Open();

                OleDbCommand DATABASE_KOMUTU;

                DATABASE_KOMUTU = new OleDbCommand
                ("INSERT INTO ceksenet(islemkodu,vade,odemetarihi,tutar,odeyecek,odemeturu) values('" + textBox5.Text + "','" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + comboBox2.Text + "')", DATABASE_BAGLAN);
                DATABASE_KOMUTU.ExecuteNonQuery();
                DATABASE_BAGLAN.Close();
            }
            catch (Exception HATA)
            {
                MessageBox.Show("OFİS DATABASE HATASI.!", HATA.Message);
            }
            MessageBox.Show(" ÇEK SENET Kaydı Başarılı.!");
            listele();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OleDbCommand kmt;
            DATABASE_BAGLAN.Open();
            kmt = new OleDbCommand("Delete from ceksenet where islemkodu = '" + dataGridView1.CurrentRow.Cells[1].Value.ToString() + "'", DATABASE_BAGLAN);
            kmt.ExecuteNonQuery();
            kmt.Dispose();
            DATABASE_BAGLAN.Close();
            MessageBox.Show("İşleminiz başarılı");
            listele();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "")
            {
                MessageBox.Show("HATALI GİRİŞ VEYA BOŞ ALAN.!");
                return;
            }
            OleDbCommand DATABASE_KOMUT;
            DATABASE_BAGLAN.Open();
            DATABASE_KOMUT = new OleDbCommand("UPDATE ceksenet set islemkodu='" + textBox5.Text + "',vade='" + textBox1.Text + "',odemetarihi='" + textBox2.Text + "',tutar='" + textBox3.Text + "',odeyecek='" + textBox4.Text + "',odemeturu='" + comboBox2.Text + "'where islemkodu='" + dataGridView1.CurrentRow.Cells[1].Value.ToString() + "'", DATABASE_BAGLAN);
            DATABASE_KOMUT.ExecuteNonQuery();
            DATABASE_BAGLAN.Close();
            MessageBox.Show("İşleminiz başarılı");
            listele();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            comboBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();

        }
        private void button4_Click(object sender, EventArgs e)
        {
            Form2 frm2 = new Form2();
            frm2.Show();
            this.Hide();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            comboBox2.Text = "";

            DATA_TABLOSU.Clear();
            DATABASE_BAGLAN.Open();
            OleDbDataAdapter adtr = new OleDbDataAdapter
          ("select * from ceksenet where islemkodu='" + textBox8.Text + "'", DATABASE_BAGLAN);
            adtr.Fill(DATA_TABLOSU);
            dataGridView1.DataSource = DATA_TABLOSU;
            dataGridView1.Columns[0].Visible = false;
            DATABASE_BAGLAN.Close();
        }
    }
}
