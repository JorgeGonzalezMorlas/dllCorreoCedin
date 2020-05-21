using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string mailDirection = @"direccionTest@test.com";
            string mailSubject = "Subject test";

            string mailContent;
            mailContent = "Mensaje de prueba";
            mailContent += "<br>" + "Otra línea";
            mailContent += "<br>" + "Otra línea";

            dllCorreos.Class1 dllc = new dllCorreos.Class1();
            dllc.SendEmailWithOutlook(mailDirection, mailSubject, mailContent);
        }
    }
}
