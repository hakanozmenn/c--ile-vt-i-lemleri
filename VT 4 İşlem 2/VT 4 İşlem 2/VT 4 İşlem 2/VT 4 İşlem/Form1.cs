using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;//Access için gerekli
using System.Drawing.Imaging;//resim işlemleri için gerekli
using System.IO;//dosya işlemleri için gerekli


namespace VT_4_İşlem
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        OleDbConnection baglantı;
        OleDbCommand komut;
        OleDbDataAdapter da;
        DataSet ds = new DataSet();

        string baglantıTeknolojisi;

        private void Form1_Load(object sender, EventArgs e)
        {
            baglantıTeknolojisi = "Provider = Microsoft.jet.Oledb.4.0;Data Source =vt4islem.mdb";
            baglantı = new OleDbConnection(baglantıTeknolojisi);
            string sqlKayıtlar = "SELECT * FROM Kisi";
            da = new OleDbDataAdapter(sqlKayıtlar, baglantı);
            baglantı.Open();
            da.Fill(ds, "Kisi");
            baglantı.Close();
            dataGridView1.DataSource = ds.Tables["Kisi"];
            textBox1.Text = ds.Tables[0].Rows[0]["ad"].ToString();
            textBox2.Text = ds.Tables[0].Rows[0]["soyad"].ToString();
            textBox3.Text = ds.Tables[0].Rows[0]["adres"].ToString();
            textBox4.Text = ds.Tables[0].Rows[0]["telefon"].ToString();
            textBox5.Text = ds.Tables[0].Rows[0]["email"].ToString();
            textBox6.Text = ds.Tables[0].Rows[0]["IDGrup"].ToString();
            dateTimePicker1.Text = ds.Tables[0].Rows[0]["tarih"].ToString();
            textBox7.Text = ds.Tables[0].Rows[0]["IDKisi"].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {//GÜNCELLEME
            baglantıTeknolojisi = "Provider = Microsoft.jet.Oledb.4.0;Data Source =vt4islem.mdb";
            baglantı = new OleDbConnection(baglantıTeknolojisi);
            string sqlKayıtGüncelle = "UPDATE Kisi SET ad=@textBox1.Text, soyad=@textBox2.Text, adres=@textBox3.Text, telefon=@textBox4.Text, email=@textBox5.Text, IDGrup=@textBox6.Text, tarih=@dateTimePicker1.Text, resim=@pictureBox1.Image WHERE IDKisi="+textBox7.Text;
            komut = new OleDbCommand(sqlKayıtGüncelle, baglantı);
            baglantı.Open();
            komut.Parameters.Add("@textBox1.Text", OleDbType.Char, 255).Value = textBox1.Text;
            komut.Parameters.Add("@textBox2.Text", OleDbType.Char, 255).Value = textBox2.Text;
            komut.Parameters.Add("@textBox3.Text", OleDbType.Char, 255).Value = textBox3.Text;
            komut.Parameters.Add("@textBox4.Text", OleDbType.Char, 255).Value = textBox4.Text;
            komut.Parameters.Add("@textBox5.Text", OleDbType.Char, 255).Value = textBox5.Text;
            komut.Parameters.Add("@textBox6.Text", OleDbType.Char, 255).Value = textBox6.Text;
            komut.Parameters.Add("@dateTimePicker1.Text", OleDbType.Char, 255).Value = dateTimePicker1.Text;
            Image resim = pictureBox1.Image;
            try
            {
                resim = Image.FromFile(openFileDialog1.FileName);
                MemoryStream ms = new MemoryStream();
                resim.Save(ms, ImageFormat.Jpeg);
                komut.Parameters.AddWithValue("@resim", ms.ToArray());
            }
            catch 
            {
                var data = (byte[])dataGridView1.Rows[0].Cells["resim"].Value;
                var stream = new MemoryStream(data);
                pictureBox1.Image = Image.FromStream(stream);
                komut.Parameters.AddWithValue("@resim", stream.ToArray());
            }
            komut.ExecuteNonQuery();
            baglantı.Close();
        }
    }
}
