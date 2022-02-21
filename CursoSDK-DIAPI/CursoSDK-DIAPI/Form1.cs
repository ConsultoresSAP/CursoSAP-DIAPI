using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CursoSDK_DIAPI
{
    public partial class Form1 : Form
    {
        private SAP sap;
        public Form1()
        {
            InitializeComponent();
            sap = new SAP();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                sap.Conectar();
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }else
                {
                    MessageBox.Show("Conectados a " + sap.CName);
                }

            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                sap.Desconectar();
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Desconectados");
                }
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
