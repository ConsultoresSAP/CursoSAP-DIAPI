﻿using System;
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

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                sap.CrearSN();
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Creado SN");
                }
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                sap.EditarSN("cliente01","Cliente01@gamail.com");
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Actualizado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                sap.AddContactoSN("cliente01","Contacto3");
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Actualizado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                sap.EditContacto("cliente01", 1);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Actualizado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                sap.agregarDireccion("cliente01");
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Actualizado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                sap.CrearItem();
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Articulo creado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                sap.AgregarAlmacen("Item001");
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Articulo modificado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
