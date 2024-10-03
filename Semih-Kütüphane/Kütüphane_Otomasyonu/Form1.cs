using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kütüphane_Otomasyonu
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.KullaniciAdi!=string.Empty)
            {
                guna2TextBox1.Text = Properties.Settings.Default.KullaniciAdi;
                guna2TextBox2.Text = Properties.Settings.Default.Sifre;
                guna2ToggleSwitch1.Checked = true;
            }
           

        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string KullaniciAdi = guna2TextBox1.Text;
                int Sifre = Convert.ToInt16(guna2TextBox2.Text);
                bool sonuc = B10.Admin(KullaniciAdi, Sifre);


                if (sonuc == false)
                {
                    guna2MessageDialog1.Show("Kullanıcı adı yada şifre yanlış!!!");
                    guna2TextBox1.Clear();
                    guna2TextBox2.Clear();
                }
                else
                {
                    Form1 kapat = new Form1();
                    kapat.Close();
                    Form2 yeni = new Form2();
                    yeni.Show();
                    this.Hide();
                }
            }
            catch (Exception)
            {

                guna2MessageDialog1.Show("Satırları Boş Bırakmayınız.");
            }
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2ToggleSwitch1_CheckedChanged(object sender, EventArgs e)
        {
            if (guna2ToggleSwitch1.Checked == true)
            {
                Properties.Settings.Default.KullaniciAdi = guna2TextBox1.Text;
                Properties.Settings.Default.Sifre = guna2TextBox2.Text;
                Properties.Settings.Default.Save();

            }

        }

        private void guna2CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void guna2CheckBox1_CheckedChanged_1(object sender, EventArgs e)
        {

            
        }
    }
}
