using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Kütüphane_Otomasyonu
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        static string baglantiYolu = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=KütüphaneBilgileri.mdb";
        static OleDbConnection baglanti = new OleDbConnection(baglantiYolu);

        private void Form3_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void çIKIŞToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void mENÜToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 kapat = new Form3();
            kapat.Close();
            Form2 ac = new Form2();
            ac.Show();
            this.Hide();
        }

        private void kİTAPEKLEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Enabled = true;
            guna2Button1.Enabled = true;
            guna2Button2.Enabled = false;
            guna2TextBox2.Visible = true;
            guna2TextBox3.Visible = true;
            guna2TextBox4.Visible = true;
            guna2TextBox1.Visible = true;
            guna2TextBox6.Visible = true;
            guna2TextBox5.Visible = true;
            guna2TextBox2.Clear();
            guna2TextBox3.Clear();
            guna2TextBox4.Clear();
            guna2TextBox1.Clear();
            guna2TextBox6.Clear();
            guna2TextBox5.Clear();
        }

        public void kitapListele()
        {
            string veri = "select*from Kitap";
            OleDbDataAdapter adaptor = new OleDbDataAdapter(veri, baglanti);
            DataSet ds = new DataSet();
            adaptor.Fill(ds);
            guna2DataGridView1.DataSource = ds.Tables[0];
        }

        private void tÜMKİTAPLARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            kitapListele();
            guna2TextBox1.Enabled = true;

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            int secilen = guna2DataGridView1.SelectedCells[0].RowIndex;
            string KitapAdı = guna2DataGridView1.Rows[secilen].Cells[1].Value.ToString();
            string SayfaNo=guna2DataGridView1.Rows[secilen].Cells[2].Value.ToString();
            string Yazar= guna2DataGridView1.Rows[secilen].Cells[3].Value.ToString();
            string BasımEvi= guna2DataGridView1.Rows[secilen].Cells[4].Value.ToString();
            guna2TextBox1.Text = KitapAdı;
            guna2TextBox3.Text = SayfaNo.ToString();
            guna2TextBox2.Text = Yazar;
            guna2TextBox4.Text = BasımEvi;
       
            kitapListele();
        }

        private void kİTAPSİLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Enabled = true;
            guna2Button1.Enabled = false;
            guna2Button2.Enabled = true;
            guna2TextBox5.Enabled = true;
            guna2TextBox1.Visible = true;
            guna2TextBox2.Visible = false;
            guna2TextBox3.Visible = false;
            guna2TextBox4.Visible = false;
            guna2TextBox6.Visible = false;
            guna2MessageDialog2.Show("Silmek İsteğin Kitabın İşlem Numarasını Giriniz");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void kİTAPBİLGİLERİGÜNCELLEToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
            try
            {
                string KitapAdı = guna2TextBox1.Text;
                int SayfaNo = Convert.ToInt32(guna2TextBox3.Text);
                string Yazar = guna2TextBox2.Text;
                string BasımEvi = guna2TextBox4.Text;
                string RafNo = guna2TextBox6.Text;
                int İslemNo = Convert.ToInt32(guna2TextBox5.Text);
                B10.KitapGuncelle(KitapAdı, SayfaNo, Yazar, BasımEvi,RafNo,İslemNo);
                guna2MessageDialog2.Show("Seçilen Kitap Başarıyla Güncellendi");
              
                guna2TextBox1.Clear();
                guna2TextBox3.Clear();
                guna2TextBox2.Clear();
                guna2TextBox4.Clear();
                guna2TextBox6.Clear();
                kitapListele();
            }
            catch (Exception )
            {
                guna2MessageDialog1.Show("Lütfen Satırları Boş Bıkramayınız.");
              
            }
            guna2TextBox5.Enabled = false;
        }

        private void kİTAPARAToolStripMenuItem_Click(object sender, EventArgs e)
        {

            guna2TextBox1.Enabled = true;
            guna2Button1.Enabled = false;
                guna2Button2.Enabled = false;
               
                guna2TextBox3.Visible = false;
                guna2TextBox2.Visible = false;
                guna2TextBox4.Visible = false;
            guna2TextBox6.Visible = false;
                guna2MessageDialog2.Show("Aramak İsteğiniz Kitabın İsmini Giriniz");   
        }

        private void button3_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            string veri = "select * from Kitap where KitapAdı like '%"+guna2TextBox1.Text+"%'";
            OleDbCommand komut = new OleDbCommand(veri, baglanti);
            OleDbDataAdapter adaptor = new OleDbDataAdapter(komut);
            DataSet DS = new DataSet();
            adaptor.Fill(DS);
            guna2DataGridView1.DataSource = DS.Tables[0];
            baglanti.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string KitapAdı = guna2TextBox1.Text;
                int SayfaNo = Convert.ToInt32(guna2TextBox3.Text);
                string Yazar = guna2TextBox2.Text;
                string BasımEvi = guna2TextBox4.Text;
                string RafNo = guna2TextBox6.Text;
                int İslemNo = Convert.ToInt32(guna2TextBox5.Text);
                guna2MessageDialog2.Show("Kitap Başarıyla Eklendi");
                B10.KitapEkle(KitapAdı, SayfaNo, Yazar, BasımEvi,RafNo,İslemNo);
                guna2Button1.Enabled = false;
                kitapListele();
                guna2TextBox1.Clear();
                guna2TextBox2.Clear();
                guna2TextBox3.Clear();
                guna2TextBox4.Clear();
                guna2TextBox6.Clear();
            }
            catch (Exception)
            {
                guna2MessageDialog2.Show("Boş Satır Bırakmayınız");
               
            }
           
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            DialogResult silinsinmi = guna2MessageDialog3.Show("Silmek İstediginden Eminmisin");
            if (silinsinmi == DialogResult.Yes)
            {
                string İslemNo = guna2TextBox5.Text;
                B10.KitapSil(İslemNo);
                guna2TextBox5.Clear();
                kitapListele();
            }
            guna2TextBox5.Visible = true;
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
          
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.Show();
            this.Close();
        }

        private void güncelleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Text = guna2DataGridView1.CurrentRow.Cells["KitapAdı"].Value.ToString();
            guna2TextBox3.Text = guna2DataGridView1.CurrentRow.Cells["SayfaNo"].Value.ToString();
            guna2TextBox2.Text = guna2DataGridView1.CurrentRow.Cells["Yazar"].Value.ToString();
            guna2TextBox4.Text = guna2DataGridView1.CurrentRow.Cells["BasımEvi"].Value.ToString();
            guna2TextBox6.Text = guna2DataGridView1.CurrentRow.Cells["RafNo"].Value.ToString();
            guna2TextBox5.Text = guna2DataGridView1.CurrentRow.Cells["İslemNo"].Value.ToString();

        }

        private void guna2DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void yazdırToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Text = guna2DataGridView1.CurrentRow.Cells["KitapAdı"].Value.ToString();
            guna2TextBox5.Text = guna2DataGridView1.CurrentRow.Cells["İslemNo"].Value.ToString();
        }

        private void guna2PictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void guna2TextBox6_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                baglanti.Open();
                string veri = "select * from Kitap where KitapAdı like '%" + guna2TextBox1.Text + "%'";
                OleDbCommand komut = new OleDbCommand(veri, baglanti);
                OleDbDataAdapter adaptor = new OleDbDataAdapter(komut);
                DataSet DS = new DataSet();
                adaptor.Fill(DS);
                guna2DataGridView1.DataSource = DS.Tables[0];
                baglanti.Close();
            }
            catch (Exception)
            {
                guna2MessageDialog2.Show("Lüften Satırı Boş Bırakmayınız.");
                throw;
            }
        }

        private void guna2Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime zaman = DateTime.Now;
            labelClock.Text = zaman.ToString();
        }

        private void guna2Button3_Click_1(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.Show();
            this.Close();
        }

        private void guna2TextBox5_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
