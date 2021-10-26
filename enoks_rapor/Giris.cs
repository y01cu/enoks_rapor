using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace enoks_rapor
{
    public partial class Giris : UserControl
    {
        public Giris()
        {
            InitializeComponent();
        }

        public static bool permission;
        public static string admin;

        private void btnGiris_Click(object sender, EventArgs e)
        {
            login();
        }

        private void login()
        {
            // kapatılmadı yalnızca görünmez hale geldi.

            if (ConfigurationManager.AppSettings["username"].ToString() == txtKullanıcıAdı.Text && ConfigurationManager.AppSettings["password"].ToString() == txtŞifre.Text)
            {
                permission = true;
                admin = txtKullanıcıAdı.Text;
                MessageBox.Show("Giriş başarılı.", "BİLGİLENDİRME", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //this.Visible = false;

            }
            else
            {
                MessageBox.Show("Giriş başarısız.", "BİLGİLENDİRME", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
