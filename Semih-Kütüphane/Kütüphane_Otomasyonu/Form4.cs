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
    public partial class Form4 : Form
    {
        
        double farkGun;
        public Form4()
        {
            InitializeComponent();
        }
        static string baglantiYolu = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=KütüphaneBilgileri.mdb";
        static OleDbConnection baglanti = new OleDbConnection(baglantiYolu);

        public void emanetListele()
        {
            string veri = "select*from Emanetler";
            OleDbDataAdapter adaptor = new OleDbDataAdapter(veri, baglanti);
            DataSet ds = new DataSet();
            adaptor.Fill(ds);
            guna2DataGridView1.DataSource = ds.Tables[0];
        }

        private void çIKIŞToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void mENÜToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 kapat = new Form4();
            kapat.Close();
            Form2 ac = new Form2();
            ac.Show();
            this.Hide();
        }

        private void tÜMEMANETLERToolStripMenuItem_Click(object sender, EventArgs e)
        {
            emanetListele();
            guna2TextBox1.Enabled = true;
        }

        private void Form4_Load(object sender, EventArgs e)
        {
           
        }

        private void eMANETKİTAPEKLEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Enabled = true;
            guna2TextBox1.Visible= true;
            guna2Button1.Enabled = true;
            guna2TextBox2.Enabled= true;
            guna2TextBox3.Enabled = true;
            guna2TextBox4.Enabled= true;
            guna2TextBox1.Visible= true;
            guna2TextBox2.Visible = true;
            guna2TextBox3.Visible = true;
            guna2TextBox4.Visible = true;
            guna2TextBox5.Enabled = true;
            label1.Visible= true;
            guna2DateTimePicker1.Visible= true;
            guna2DateTimePicker2.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void eMANETKİTAPSİLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox5.Enabled = true;
            guna2TextBox1.Enabled = true;
            guna2Button2.Enabled = true;
            guna2TextBox1.Visible = true;
            guna2TextBox1.Visible = true;
            guna2TextBox2.Visible = false;
            guna2TextBox3.Visible = false;
            guna2TextBox4.Visible = false;
            label1.Visible= false;
            guna2DateTimePicker1.Visible = false;
            guna2DateTimePicker2.Visible = false;
            guna2MessageDialog1.Show("Silmek İsteğin Emanet Kitabın İsmini Giriniz");
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int secilen = guna2DataGridView1.SelectedCells[0].RowIndex;
            string KitapAdı = guna2DataGridView1.Rows[secilen].Cells[1].Value.ToString();
            int  KitapNo=Convert.ToInt32(guna2DataGridView1.Rows[secilen].Cells[2].Value);
            string ÜyeAdı = guna2DataGridView1.Rows[secilen].Cells[3].Value.ToString();
            int ÜyeNo =Convert.ToInt32(guna2DataGridView1.Rows[secilen].Cells[4].Value);
            string AldığıTarih = guna2DataGridView1.Rows[secilen].Cells[5].Value.ToString();

            guna2TextBox1.Text = KitapAdı;
            guna2TextBox2.Text = KitapNo.ToString();
            guna2TextBox3.Text = ÜyeAdı;
            guna2TextBox4.Text = ÜyeNo.ToString();
            guna2DateTimePicker1.Text = AldığıTarih;
            emanetListele();
        }

        private void eMANETKİTAPGÜNCELLEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string KitapAdı = guna2TextBox1.Text;
                int KitapNo = Convert.ToInt32(guna2TextBox2.Text);
                string ÜyeAdı = guna2TextBox3.Text;
                int ÜyeNo = Convert.ToInt32(guna2TextBox4.Text);
                string AldığıTarih = guna2DateTimePicker1.Text;
                string TeslimTarih = guna2DateTimePicker1.Text;
                int İslemNo = Convert.ToInt32(guna2TextBox5.Text);

                B10.emanetGuncelle(KitapAdı, KitapNo, ÜyeAdı, ÜyeNo, AldığıTarih,TeslimTarih,İslemNo);
                guna2MessageDialog2.Show("Seçilen Kitap Başarıyla Güncellendi");
                guna2TextBox1.Clear();
                guna2TextBox2.Clear();
                guna2TextBox3.Clear();
                guna2TextBox4.Clear();
                guna2TextBox5.Clear();
                emanetListele();
                
            }
            catch (Exception)
            {
                guna2MessageDialog1.Show("Lütfen Satırları Boş Bıkramayınız.");
            }
            guna2TextBox5.Enabled = false;

        }

        private void eMANETKİTAPARAToolStripMenuItem_Click(object sender, EventArgs e)
        {

            guna2TextBox1.Visible = true;
            guna2Button1.Enabled = false;
            guna2Button2.Enabled = false;
    
            guna2TextBox2.Visible = false;
            guna2TextBox3.Visible = true;
            guna2TextBox1.Visible = false;
            guna2TextBox4.Visible = false;
            guna2DateTimePicker1.Visible = false;
            guna2DateTimePicker2.Visible = false;
            label1.Visible= false;
            guna2MessageDialog1.Show("Aramak İsteğiniz Üyenin İsmini Giriniz");
        }

        private void button3_Click(object sender, EventArgs e)
        {
          
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {

            try
            {
                string KitapAdı = guna2TextBox1.Text;
                int KitapNo = Convert.ToInt32(guna2TextBox2.Text);
                string ÜyeAdı = guna2TextBox3.Text;
                int ÜyeNo = Convert.ToInt32(guna2TextBox4.Text);
                string AldığıTarih = guna2DateTimePicker1.Text;
                string SonTeslimTarih = guna2DateTimePicker2.Text;
                int İslemNo = Convert.ToInt32(guna2TextBox5.Text);
                B10.emanetEkle(KitapAdı, KitapNo, ÜyeAdı, ÜyeNo, AldığıTarih, SonTeslimTarih,İslemNo);
                guna2MessageDialog1.Show("Emanet Kitaplar Başarıyla Eklendi");
            }
            catch (Exception)
            {
                guna2MessageDialog2.Show("Boşlukları Doldurunuz.");
            }
                guna2TextBox1.Enabled = false;
                emanetListele();
                guna2TextBox1.Clear();
                guna2TextBox2.Clear();
                guna2TextBox3.Clear();
                guna2TextBox4.Clear();
            guna2TextBox5.Clear();


        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            DialogResult silinsinmi = guna2MessageDialog3.Show("Silmek İstediginden Eminmisin");
            if (silinsinmi == DialogResult.Yes)
            {
                string İslemNo = guna2TextBox5.Text;
                B10.emanetSil(İslemNo);
                guna2TextBox1.Clear();
                guna2TextBox5.Clear();
            }
            emanetListele();
            guna2Button2.Enabled = false;
            guna2TextBox2.Visible = true;
            guna2TextBox3.Visible = true;
            guna2TextBox4.Visible = true;
            guna2DateTimePicker1.Visible = true;
            guna2DateTimePicker2.Visible = true;
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
           
        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.ShowDialog();
            this.Close();
        }
        public void kitapListele()
        {
            string veri = "select*from Kitap";
            OleDbDataAdapter adaptor = new OleDbDataAdapter(veri, baglanti);
            DataSet ds = new DataSet();
            adaptor.Fill(ds);
            guna2DataGridView1.DataSource = ds.Tables[0];
        }

        private void kitaplarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            kitapListele();
            guna2TextBox1.Enabled = true;
        }

        private void guna2DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
         

        }

        private void guna2Panel1_Paint(object sender, PaintEventArgs e)
        {

        }
       
        private void güncelleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (guna2DataGridView1.Rows.Count > 0)
                {


                    guna2TextBox1.Text = guna2DataGridView1.CurrentRow.Cells["KitapAdı"].Value.ToString();
                    guna2TextBox2.Text = guna2DataGridView1.CurrentRow.Cells["KitapNo"].Value.ToString();
                    guna2TextBox3.Text = guna2DataGridView1.CurrentRow.Cells["ÜyeAdı"].Value.ToString();
                    guna2TextBox4.Text = guna2DataGridView1.CurrentRow.Cells["ÜyeNo"].Value.ToString();
                    guna2DateTimePicker1.Text = guna2DataGridView1.CurrentRow.Cells["AldığıTarih"].Value.ToString();
                    guna2DateTimePicker2.Text = guna2DataGridView1.CurrentRow.Cells["SonTeslimTarih"].Value.ToString();
                }
            }
            catch (Exception)
            {
                guna2MessageDialog2.Show("Lütfen Bu Özelliği TümEmanetKitaplar Kısmında Kullanınız.");

            }
        }

        private void yazdırToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                guna2TextBox1.Text = guna2DataGridView1.CurrentRow.Cells["KitapAdı"].Value.ToString();
                guna2TextBox2.Text = guna2DataGridView1.CurrentRow.Cells["KitapNumarası"].Value.ToString();
            }
            catch (Exception)
            {
                guna2MessageDialog2.Show("Lütfen Bu Özelliği Kitap Eklenirken Kullanınız.");
            }
            
        }

        private void guna2TextBox4_TextChanged(object sender, EventArgs e)
        {

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

        private void guna2TextBox3_TextChanged(object sender, EventArgs e)
        {
            baglanti.Open();
            string veri = "select * from Emanetler where ÜyeAdı like '%" + guna2TextBox3.Text + "%'";
            OleDbCommand komut = new OleDbCommand(veri, baglanti);
            OleDbDataAdapter adaptor = new OleDbDataAdapter(komut);
            DataSet DS = new DataSet();
            adaptor.Fill(DS);
            guna2DataGridView1.DataSource = DS.Tables[0];
            baglanti.Close();
        }

        private void guna2PictureBox1_Click_1(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.Show();
            this.Close();
        }

        private void kitapİsmiYazıdırToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guna2TextBox1.Text = guna2DataGridView1.CurrentRow.Cells["KitapAdı"].Value.ToString();
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

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}
