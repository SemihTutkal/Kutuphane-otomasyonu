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
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
            timer1.Start();
        }

        static string baglantiYolu = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=KütüphaneBilgileri.mdb";
        static OleDbConnection baglanti = new OleDbConnection(baglantiYolu);

        private void çIKIŞToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void mENÜToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form5 kapat = new Form5();
            kapat.Close();
            Form2 ac = new Form2();
            ac.Show();
            this.Hide();
        }

        private void Form5_Load(object sender, EventArgs e)
        {

        }
        public void üyeleriListele()
        {
            string veri = "select*from Üyeler";
            OleDbDataAdapter adaptor = new OleDbDataAdapter(veri, baglanti);
            DataSet ds = new DataSet();
            adaptor.Fill(ds);
            guna2DataGridView1.DataSource = ds.Tables[0];
        }
        private void tÜMÜYELERToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Enabled = true;
            üyeleriListele();
        }

        private void üYEEKLEToolStripMenuItem_Click(object sender, EventArgs e)
        {
          guna2TextBox1.Enabled=true;
            guna2TextBox3.Visible = true;
            guna2TextBox4.Visible = true;
            guna2Button1.Enabled = true;
            guna2TextBox3.Enabled = true;
            guna2TextBox4.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ÜyeAdSoyad = guna2TextBox1.Text;
            string Sınıf = guna2TextBox3.Text;
            int OkulNo = Convert.ToInt32(guna2TextBox4.Text);

            B10.üyeEkle(ÜyeAdSoyad, Sınıf, OkulNo);
            guna2MessageDialog1.Show("Üye Başarıyla Eklendi");
            guna2Button1.Enabled = false;
            üyeleriListele();
            guna2TextBox1.Clear();
            guna2TextBox3.Clear();
            guna2TextBox4.Clear();

        }

        private void üYESİLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Enabled = true;
            guna2Button2.Enabled = true;

            guna2TextBox3.Visible = false;
            guna2TextBox4.Visible = false;
            guna2MessageDialog1.Show("Silmek İsteğin Üyenin İsmini Gir");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string ÜyeAdı = guna2TextBox1.Text;
            B10.üyeSil(ÜyeAdı);
            guna2MessageDialog1.Show("İstenilen Üye Başarıyla Silindi");
            guna2TextBox1.Clear();
            üyeleriListele();
            guna2Button2.Enabled = false;
  

            guna2TextBox3.Visible = true;
            guna2TextBox4.Visible = true;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int secilen = guna2DataGridView1.SelectedCells[0].RowIndex;
            string ÜyeAdSoyad = guna2DataGridView1.Rows[secilen].Cells[1].Value.ToString();
            string Sınıf = guna2DataGridView1.Rows[secilen].Cells[3].Value.ToString();
            int OkulNo = Convert.ToInt32(guna2DataGridView1.Rows[secilen].Cells[4].Value);

            guna2TextBox1.Text = ÜyeAdSoyad;
            guna2TextBox3.Text = Sınıf;
            guna2TextBox4.Text = OkulNo.ToString();
            üyeleriListele();
        }

        private void üYEGÜNCELLEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                baglanti.Open();
                string ÜyeAdSoyad = guna2TextBox1.Text;
                string Sınıf = guna2TextBox3.Text;
                int OkulNo = Convert.ToInt32(guna2TextBox4.Text);
                B10.üyeGuncelle(ÜyeAdSoyad, Sınıf, OkulNo);
                guna2MessageDialog2.Show("Seçilen Kitap Başarıyla Güncellendi");
                guna2TextBox1.Clear();
                guna2TextBox3.Clear();
                guna2TextBox4.Clear();
                üyeleriListele();
                baglanti.Close();

            }
            catch (Exception)
            {
                guna2MessageDialog1.Show("Lütfen Satırları Boş Bıkramayınız.");
            }
          guna2TextBox1.Enabled= false;


        }

        private void üYEARAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Enabled = true;
            guna2Button1.Enabled = false;
            guna2Button2.Enabled = false;
            guna2TextBox3.Visible = false;
            guna2TextBox4.Visible = false;
            guna2MessageDialog1.Show("Aramak İsteğiniz Üyenin İsmini Ve Soyismini Giriniz");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            string veri = "select * from Üyeler where ÜyeAdı like '%" + guna2TextBox1.Text + "%'";
            OleDbCommand komut = new OleDbCommand(veri, baglanti);
            OleDbDataAdapter adaptor = new OleDbDataAdapter(komut);
            DataSet DS = new DataSet();
            adaptor.Fill(DS);
            guna2DataGridView1.DataSource = DS.Tables[0];
            baglanti.Close();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string ÜyeAdSoyad = guna2TextBox1.Text;
                string Sınıf = guna2TextBox3.Text;
                int OkulNo = Convert.ToInt32(guna2TextBox4.Text);

                B10.üyeEkle(ÜyeAdSoyad, Sınıf, OkulNo);
                guna2MessageDialog1.Show("Üye Başarıyla Kaydedildi");
            }
            catch (Exception)
            {
                guna2MessageDialog2.Show("Lütfen Satırları Boş Bırakmayınız.");
               
            }
        
            üyeleriListele();
            guna2TextBox1.Clear();
            guna2TextBox3.Clear();
            guna2TextBox4.Clear();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            
            
            
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            DialogResult silinsinmi = guna2MessageDialog3.Show("Silmek İstediginden Eminmisin");
            if (silinsinmi == DialogResult.Yes)
            {
                string ÜyeAdı = guna2TextBox1.Text;
                B10.üyeSil(ÜyeAdı);
                guna2TextBox1.Clear();
            }
                
            üyeleriListele();
            guna2Button2.Enabled = false;
            guna2TextBox3.Visible = true;
            guna2TextBox4.Visible = true;
        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.Show();
            this.Close();
        }

        private void güncelleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Text = guna2DataGridView1.CurrentRow.Cells["ÜyeAdSoyad"].Value.ToString();
            guna2TextBox3.Text = guna2DataGridView1.CurrentRow.Cells["Sınıf"].Value.ToString();
            guna2TextBox4.Text = guna2DataGridView1.CurrentRow.Cells["OkulNo"].Value.ToString();
        }

        private void guna2DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void yazdırToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Text = guna2DataGridView1.CurrentRow.Cells["ÜyeAdSoyad"].Value.ToString();
        }

        private void guna2PictureBox2_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    DialogResult dialog = new DialogResult();
            //    dialog = guna2MessageDialog3.Show("Bu işlem, veri yoğunluğuna göre uzun sürebilir. Devam etmek istiyor musunuz?", "EXCEL'E AKTARMA");
            //    if (dialog == DialogResult.Yes)
            //    {
            //        Microsoft.Office.Interop.Excel.Application uyg = new Microsoft.Office.Interop.Excel.Application();
            //        uyg.Visible = true;
            //        Microsoft.Office.Interop.Excel.Workbook kitap = uyg.Workbooks.Add(System.Reflection.Missing.Value);
            //        Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
            //        for (int i = 0; i < guna2DataGridView1.Columns.Count; i++)
            //        {
            //            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[1, i + 1];
            //            myRange.Value2 = guna2DataGridView1.Columns[i].HeaderText;
            //        }

            //        for (int i = 0; i < guna2DataGridView1.Columns.Count; i++)
            //        {
            //            for (int j = 0; j < guna2DataGridView1.Rows.Count; j++)
            //            {
            //                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
            //                myRange.Value2 = guna2DataGridView1[i, j].Value;
            //            }
            //        }
            //    }
            //    else
            //    {
            //        guna2MessageDialog1.Show("İŞLEM İPTAL EDİLDİ.", "İşlem Sonucu");
            //    }
            //}
            //catch (Exception)
            //{

            //    guna2MessageDialog1.Show("İŞLEM TAMAMLANMADAN EXCEL PENCERESİNİ KAPATTINIZ.", "HATA");
            //}
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {

        }

        private void guna2Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime zaman = DateTime.Now;
            labelClock.Text = zaman.ToString();
        }

        private void labelClock_Click(object sender, EventArgs e)
        {
            
        }

        private void guna2TextBox4_TextChanged(object sender, EventArgs e)
        {

            baglanti.Open();
            string veri = "select * from Üyeler where OkulNo like '%" + guna2TextBox4.Text + "%'";
            OleDbCommand komut = new OleDbCommand(veri, baglanti);
            OleDbDataAdapter adaptor = new OleDbDataAdapter(komut);
            DataSet DS = new DataSet();
            adaptor.Fill(DS);
            guna2DataGridView1.DataSource = DS.Tables[0];
            baglanti.Close();
        }

        private void guna2Button3_Click_1(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.Show();
            this.Close();
        }
    }
}
