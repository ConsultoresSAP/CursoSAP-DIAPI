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

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                sap.CrearPedido(out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Pedido #"+DocEntry+" creado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                int DocEntry = 564;
                sap.AgregarLineaPedido(DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Pedido #" + DocEntry.ToString() + " fue actualizado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                sap.AgregarPedidoTipoServicio(out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Pedido #" + DocEntry + " creado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                sap.CrearEntrega(out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Entrega #" + DocEntry + " creado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                sap.CrearDevolucion(out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Devolucion #" + DocEntry + " creado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                sap.CrearSalida(out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Salida #" + DocEntry + " creada con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                sap.CrearFacturaConDocumentoBase(out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Factura #" + DocEntry + " creada con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                sap.CrearPedidoEnBaseABorrador(out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Pedido #" + DocEntry + " creada con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                sap.CrearTransferencia(out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Transferencia #" + DocEntry + " creada con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                int DocEntryFac = 221;
                string DocEntryPago = "";
                sap.CrearPago(DocEntryFac,out DocEntryPago);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Pago #" + DocEntryPago + " creado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                int DocEntryFac = 221;
                string Datos = "";
                sap.Record(DocEntryFac, out Datos);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show(Datos);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                string DocNumFac = "221";
                string DocEntryPago = "";
                sap.CrearPago(DocNumFac, out DocEntryPago);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Pago #" + DocEntryPago + " creado con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                string DocNumPedido = "478";
                sap.CrearFacturaConDocumentoBase(DocNumPedido, out DocEntry);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Factura #" + DocEntry + " creada con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            try
            {
                sap.ActualizarListaDePrecios(2,1, "009-001-001-000001");
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + sap.Error);
                }
                else
                {
                    MessageBox.Show("Lista actualizada con exito");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
