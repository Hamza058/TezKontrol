using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TezKontrol
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<string> metin = new List<string>();

        private void button1_Click(object sender, EventArgs e)
        {
            Bul();
            onsoz();
            grs();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { ValidateNames = true, Multiselect = false, Filter = "Word Dosyası |*.docx" })//docx dosyasını tanımlayıp sisteme yüklenmesini sağlıyoruz.
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    object readOnly = false;
                    object visible = true;
                    object save = false;
                    object fileName = ofd.FileName;
                    object newTemplate = false;
                    object docType = 0;
                    object missing = Type.Missing;
                    Microsoft.Office.Interop.Word._Document document;
                    Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application()
                    {
                        Visible = false
                    };
                    document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref visible, ref missing, ref missing, ref missing, ref missing);//Word dosyasını açıyoruz.
                    document.ActiveWindow.Selection.WholeStory();
                    document.ActiveWindow.Selection.Copy();
                    IDataObject dataObject = Clipboard.GetDataObject();
                    richTextBox1.Rtf = dataObject.GetData(DataFormats.Rtf).ToString();//word dosyasındaki bilgileri richTextBox a atama yapıyoruz.
                    application.Quit(ref missing, ref missing, ref missing);
                    metin.Add(richTextBox1.Text);//richTextBox daki bilgileri metin değişkenine ekliyoruz
                }
            }
        }
        public void Bul()
        {
            foreach (string item in metin)//metin değişkenindeki bilgileri item e aktarıyoruz
            {
                int sayac = 0;
                for (int i = 0; i < item.Length; i++)
                {
                    if (item[i] == '“')
                    {
                        sayac++;
                    }
                    if (item[i] == '”')
                    {
                        sayac++;
                    }
                    label1.Visible = true;
                    label4.Visible = true;
                    int sonuc = sayac / 2;
                    if (sonuc > 50)
                        label4.Text = "Çift tırnak arasındaki\nkelime sayısı > 50";
                    else
                        label4.Text = "Çift tırnak arasındaki\nkelime sayısı < 50";
                    label1.Text = "Çift tırnak arasındaki\nkelime sayısı: " + sonuc.ToString();
                }
            }
        }
        public void onsoz()
        {
            int sonuc;
            int sonuc2 = 0;
            int sonuc3 = 0;
            label2.Visible = true;
            foreach (string item in metin)//metin değişkenindeki bilgileri item e aktarıyoruz
            {
                sonuc = item.IndexOf("ÖNSÖZ");//Metinde ÖNSÖZ kelimesini aratıyoruz. Varsa index ini verecek yoksa -1 değeri döndürecek.
                if (sonuc > 0)//Metin de ÖNSÖZ var mı yok mu kontrol ediyoruz 
                {
                    sonuc2 = item.IndexOf(".");
                    sonuc3 = item.IndexOf("teşekkür");
                    if (sonuc2 > sonuc3)
                    {
                        label2.Text = "Önsöz beyanemesinin ilk paragrafında\nteşekkür ibaresi yer alıyor.";
                        break;
                    }
                    else if (sonuc3 > sonuc2)
                    {
                        label2.Text = "Önsöz beyanemesinin ilk paragrafında\nteşekkür ibaresi yer almıyor.";
                        break;
                    }
                }
            }
        }
        public void grs()
        {
            string a = "GİRİŞ";
            string b = "organizasyon";
            string c = "kapsam";
            string d = "GENEL YAZIM KURALLARI";
            label3.Visible = true;

            int sonuc;

            foreach (string item in metin)//metin değişkenindeki bilgileri item e aktarıyoruz
            {
                sonuc = item.IndexOf(a);
                if (sonuc > 0)
                {
                    if (item.IndexOf(d) > item.IndexOf(b) && item.IndexOf(d) > item.IndexOf(c))
                    {
                        label3.Text = "Giriş bölümünde\norganizasyon ve kapsam bulunmaktadır.";
                    }
                    else
                        label3.Text = "Giriş bölümünde\norganizasyon ve kapsam bulunmamaktadır.";
                }
            }
        }
    }
}
