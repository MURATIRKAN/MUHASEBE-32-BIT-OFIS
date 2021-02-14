using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
namespace OnMuhasebeOtomasyonu
{
    public partial class kullanici : Form
    {
        public kullanici()
        {
            InitializeComponent();
        }
        OleDbConnection baglan = new OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" + Application.StartupPath + "\\datam.accdb");
        DataTable tablo = new DataTable();

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
          ("select * from kullanici", baglan);
            adtr.Fill(tablo);
            dataGridView1.DataSource = tablo;
            dataGridView1.Columns[0].Visible = false;
            baglan.Close();
        }
        private void kullanici_Load(object sender, EventArgs e)
        {
            listele();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

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
                    listele();
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

        private void button3_Click(object sender, EventArgs e)
        {

            OleDbCommand kmt;
            baglan.Open();
            kmt = new OleDbCommand("Delete from kullanici where tc = '" + dataGridView1.CurrentRow.Cells[1].Value.ToString() + "'", baglan);
            kmt.ExecuteNonQuery();
            kmt.Dispose();
            baglan.Close();
            MessageBox.Show("İşleminiz başarılı");
            listele();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OleDbCommand kmt;
            baglan.Open();
            kmt = new OleDbCommand("UPDATE kullanici SET tc='" + textBox1.Text + "',adsoyad='" + textBox2.Text + "',tel='" + textBox3.Text + "',k_adi='" + textBox4.Text + "',sifre='" + textBox5.Text + "',gizlisoru='" + comboBox1.Text + "',yaniti='" + textBox7.Text + "'where tc='" + dataGridView1.CurrentRow.Cells[1].Value.ToString() + "'", baglan);
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
            textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            textBox7.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form2 FORMANA = new Form2();
            FORMANA.Show();
            this.Hide();
        }
    }
}
