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
    public partial class Form2 : Form
    {

        private int ImageNumber = 1;

        static string baglantiYolu = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=KütüphaneBilgileri.mdb";
        static OleDbConnection baglanti = new OleDbConnection(baglantiYolu);

        public Form2()
        {
           
           
            InitializeComponent();
        }

        private void çIKIŞToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void hAKKIMIZDAToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void hESAPMAKİNESİToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void Form2_Load(object sender, EventArgs e)
        {
           
        }

        public void kitapListele(string veri)
        {
            OleDbDataAdapter adaptor = new OleDbDataAdapter(veri,baglanti);
            DataSet ds = new DataSet();
            adaptor.Fill(ds);
            guna2DataGridView1.DataSource = ds.Tables[0];
        }
        private void button2_Click(object sender, EventArgs e)
        {
          
            guna2DataGridView1.Visible = true;
            kitapListele("select*from Kitap");
        }

        public void emanetListele(string veri)
        {
            OleDbDataAdapter adaptor = new OleDbDataAdapter(veri, baglanti);
            DataSet ds = new DataSet();
            adaptor.Fill(ds);
            guna2DataGridView1.DataSource = ds.Tables[0];
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            guna2DataGridView1.Visible = true;
            emanetListele("select*from Emanetler");
        }

        public void üyeleriListele(string veri)
        {
            OleDbDataAdapter adaptor = new OleDbDataAdapter(veri, baglanti);
            DataSet ds = new DataSet();
            adaptor.Fill(ds);
            guna2DataGridView1.DataSource = ds.Tables[0];
        }

        private void button5_Click(object sender, EventArgs e)
        {
         
            guna2DataGridView1.Visible = true;
            üyeleriListele("select*from Üyeler");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void kİTAPARAToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
         
            guna2DataGridView1.Visible = true;
            emanetListele("select*from Emanetler");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form2 kapat = new Form2();
            kapat.Close();
            Form5 yeni = new Form5();
            yeni.Show();
            this.Hide();
        }

        private void eRİŞİMKOLAYLIĞIToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            Form2 kapat = new Form2();
            kapat.Close();
            Form3 yeni = new Form3();
            yeni.Show();
            this.Hide();
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            Form2 kapat = new Form2();
            kapat.Close();
            Form4 yeni = new Form4();
            yeni.Show();
            this.Hide();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Form2 kapat = new Form2();
            kapat.Close();
            Form5 yeni = new Form5();
            yeni.Show();
            this.Hide();
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            try
            {
                guna2Panel1.Visible= false;
                pictureBox1.Visible = false;
            }
            catch (Exception)
            {
                pictureBox1.Visible = true;
                throw;
            }

                //pictureBox1.Visible = false;
                guna2DataGridView1.Visible = true;
            kitapListele("select*from Kitap");
        }

        private void guna2Button7_Click(object sender, EventArgs e)
        {
            try
            {
                guna2Panel1.Visible = false;
                pictureBox1.Visible = false;
            }
            catch (Exception)
            {
                pictureBox1.Visible = true;
                throw;
            }


            // pictureBox1.Visible = false;
            guna2DataGridView1.Visible = true;
            emanetListele("select*from Emanetler");
        }

        private void guna2Button8_Click(object sender, EventArgs e)
        {
            try
            {
                guna2Panel1.Visible = false;
                pictureBox1.Visible = false;
            }
            catch (Exception)
            {
                pictureBox1.Visible = true;
                throw;
            }



            // pictureBox1.Visible = false;
            labelClock.Visible = true;
            guna2DataGridView1.Visible = true;
            üyeleriListele("select*from Üyeler");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button3_Click_1(object sender, EventArgs e)
        {
            Form2 kapat = new Form2();
            kapat.Close();
            Form5 yeni = new Form5();
            yeni.Show();
            this.Hide();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime zaman = DateTime.Now;
            labelClock.Text = zaman.ToString();
        }

        private void labelClock_Click(object sender, EventArgs e)
        {

        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {

        }
       

        private void timer2_Tick(object sender, EventArgs e)
        {
            Slider();

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {


        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.ShowDialog();
            this.Close();
        }

        private void guna2DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
      

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {

        }
        private void Slider()
        {
            if (ImageNumber == 4)
            {
                ImageNumber = 1;
            }
            pictureBox1.ImageLocation = string.Format(@"Images\{0}.jpg", ImageNumber);
            ImageNumber++;
        }

        private void pictureBox1_Click_2(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.ShowDialog();
            this.Close();
        }

        private void guna2CirclePictureBox4_Click(object sender, EventArgs e)
        {
            Form2 ana = new Form2();
            ana.Show();
            this.Close();
        }
    }
}
