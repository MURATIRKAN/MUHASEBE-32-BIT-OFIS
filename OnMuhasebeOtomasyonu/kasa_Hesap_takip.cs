using System;
using System.Data;
using System.Windows.Forms;
#pragma warning disable CS0105 // The using directive for 'System.Data' appeared previously in this namespace
#pragma warning restore CS0105 // The using directive for 'System.Data' appeared previously in this namespace
using System.Data.OleDb;//access veritabanı işlemlerimiz için ekledi
namespace OnMuhasebeOtomasyonu
{
    public partial class kasa_Hesap_takip : Form
    {
        public kasa_Hesap_takip()
        {
            InitializeComponent();
        }

        OleDbConnection baglan = new OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\datam.accdb");
        DataTable tablo = new DataTable();
        int SAY;
        public void listele()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox7.Text = "";
            comboBox1.Text = "";


            tablo.Clear();
            baglan.Open();
            OleDbDataAdapter adtr = new OleDbDataAdapter
          ("select * from kasa", baglan);
            adtr.Fill(tablo);
            dataGridView1.DataSource = tablo;
            dataGridView1.Columns[0].Visible = false;
            baglan.Close();
        }

        private void kasa_Hesap_takip_Load(object sender, EventArgs e)
        {
            listele();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SAY = tablo.Rows.Count;
            if (SAY > 5)
            {
                MessageBox.Show("SINIRLI KAYIT.!LÜFTEN TAM SÜRÜMÜ İSTEYİNİZ.!", "SINIRLI KULLANIM BİLGİSİ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (baglan.State == ConnectionState.Open)
            {

                baglan.Close();
            }
            baglan.Open();

            OleDbCommand kmt;

            kmt = new OleDbCommand
            ("INSERT INTO kasa(tarih,cari,aciklama,tip,masraf,giren,cikan,islemkodu) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + comboBox1.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "')", baglan);
            kmt.ExecuteNonQuery();
            baglan.Close();
            MessageBox.Show("Kayıt Başarılı");
            listele();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox1.Text = dateTimePicker1.Value.ToShortDateString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("HATALI BOŞ ALAN.!");
                return;
            }
            OleDbCommand kmt;
            baglan.Open();
            kmt = new OleDbCommand("Delete from kasa where islemkodu = '" + dataGridView1.CurrentRow.Cells[8].Value.ToString() + "'", baglan);
            kmt.ExecuteNonQuery();
            kmt.Dispose();
            baglan.Close();
            MessageBox.Show("İşleminiz başarılı");
            listele();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OleDbCommand kmt;
            baglan.Open();
            kmt = new OleDbCommand("UPDATE kasa SET tarih='" + textBox1.Text + "',cari='" + textBox2.Text + "',aciklama='" + textBox3.Text + "',tip='" + comboBox1.Text + "',masraf='" + textBox4.Text + "',giren='" + textBox5.Text + "',cikan='" + textBox6.Text + "',islemkodu='" + textBox7.Text + "'where islemkodu='" + dataGridView1.CurrentRow.Cells[8].Value.ToString() + "'", baglan);
            kmt.ExecuteNonQuery();
            baglan.Close();
            MessageBox.Show("İşleminiz başarılı");
            listele();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            textBox7.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
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
            textBox5.Text = "";
            textBox6.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox7.Text = "";
            comboBox1.Text = "";


            tablo.Clear();
            baglan.Open();
            OleDbDataAdapter adtr = new OleDbDataAdapter
          ("select * from kasa where islemkodu = '" + textBox8.Text + "'", baglan);
            adtr.Fill(tablo);
            dataGridView1.DataSource = tablo;
            dataGridView1.Columns[0].Visible = false;
            baglan.Close();
        }
    }
}
