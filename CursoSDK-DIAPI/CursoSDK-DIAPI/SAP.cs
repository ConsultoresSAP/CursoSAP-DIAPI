using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace CursoSDK_DIAPI
{
  

    class SAP
    {
        private Company oCom;
        public string Error = "";
        public string CName = "";

        public SAP()
        {
            this.oCom = new Company();
        }

        public void Conectar()
        {
            try
            {
                this.oCom.Server = "LABORATORIO";
                this.oCom.DbServerType = BoDataServerTypes.dst_MSSQL2014;
                this.oCom.CompanyDB = "SBODemoGT";
                this.oCom.UserName = "manager";
                this.oCom.Password = "manager";

                if (this.oCom == null)
                {
                    this.oCom = new Company();
                }

                if (!this.oCom.Connected)
                {
                    int ErrorCode = this.oCom.Connect();

                    if (ErrorCode != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription() + " (" + ErrorCode.ToString() + " )";
                    }else
                    {
                        this.CName = this.oCom.CompanyName;
                    }
                }


            }catch(Exception e)
            {
                this.Error = e.Message;
            }
        }


        public void Desconectar()
        {
            try
            {
                this.Error = "";
                if (this.oCom != null)
                {
                    if (this.oCom.Connected)
                    {
                        this.oCom.Disconnect();
                        this.CName = "";
                    }
                }

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (this.oCom != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.oCom);
                    this.oCom = null;
                    GC.Collect();
                }
            }
        }


        public void CrearSN()
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                this.Error = "";
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                oSN.CardCode = "Cliente01";
                oSN.CardName = "Cliente de prueba numero 1";
                oSN.CardType = BoCardTypes.cCustomer;
                oSN.FederalTaxID = "123456789101";

                if (oSN.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription() + " ( " + this.oCom.GetLastErrorCode() + " )";
                }

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        public void EditarSN(string CardCode, string Correo)
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                this.Error = "";
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (oSN.GetByKey(CardCode))
                {
                    oSN.EmailAddress = Correo;
                    if (oSN.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription() + " ( " + this.oCom.GetLastErrorCode() + " )";
                    }
                }else
                {
                    this.Error = "SN no existe";
                }


            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }



        public void AddContactoSN(string CardCode,string Name)
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (oSN.GetByKey(CardCode))
                {
                    if (oSN.ContactEmployees.Count > 1)
                    {
                        oSN.ContactEmployees.Add();
                    }else
                    {
                        if (oSN.ContactEmployees.Name != "")
                        {
                            oSN.ContactEmployees.Add();
                        }
                    }

                    oSN.ContactEmployees.Name = Name;
                    oSN.ContactEmployees.FirstName = "Pedro";
                    oSN.ContactEmployees.MiddleName = "Juan";
                    oSN.ContactEmployees.LastName = "Perez";
                    oSN.ContactEmployees.Title = "SR";

                    if (oSN.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription() + " ( " + this.oCom.GetLastErrorCode() + " ) ";
                    }

                }else
                {
                    this.Error = "SN no existe";
                }

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        public void EditContacto(string CardCode,int line)
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                if (oSN.GetByKey(CardCode))
                {

                    oSN.ContactEmployees.SetCurrentLine(line);

                    oSN.ContactEmployees.FirstName = "Francisco";

                    if (oSN.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription() + " ( " + this.oCom.GetLastErrorCode() + " ) ";
                    }

                }
                else
                {
                    this.Error = "SN no existe";
                }

            }
            catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        public void agregarDireccion(string CardCode)
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                if (oSN.GetByKey(CardCode))
                {
                    if (oSN.Addresses.Count > 1)
                    {
                        oSN.Addresses.Add();
                    }
                    else
                    {
                        if (oSN.Addresses.AddressName != "")
                        {
                            oSN.Addresses.Add();
                        }
                    }

                    oSN.Addresses.AddressName = "Direccion 3";
                    oSN.Addresses.AddressType = BoAddressType.bo_BillTo;

                    if (oSN.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription() + " ( " + this.oCom.GetLastErrorCode() + " ) ";
                    }

                }
                else
                {
                    this.Error = "SN no existe";
                }
            }
            catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        public void CrearItem()
        {
            SAPbobsCOM.Items oItem = null;
            try
            {
                this.Error = "";
                oItem = (SAPbobsCOM.Items)this.oCom.GetBusinessObject(BoObjectTypes.oItems);

                oItem.ItemCode = "Item001";
                oItem.ItemName = "Celular";

                oItem.WhsInfo.WarehouseCode = "01";
                oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES;

                if (oItem.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oItem != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                    oItem = null;
                }
            }
        }


        public void AgregarAlmacen(string ItemCode)
        {
            SAPbobsCOM.Items oItem = null;
            try
            {
                this.Error = "";
                oItem = (SAPbobsCOM.Items)this.oCom.GetBusinessObject(BoObjectTypes.oItems);

                if (oItem.GetByKey(ItemCode))
                {
                    oItem.WhsInfo.Add();
                    oItem.WhsInfo.WarehouseCode = "5";

                    if (oItem.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }

                }
                else
                {
                    this.Error = "Articulo no existe";
                }

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oItem != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                    oItem = null;
                }
            }
        }

        #region Documents

        public void CrearPedido(out string DocEntry)
        {
            SAPbobsCOM.Documents oPedido = null;
            DocEntry = "";
            try
            {
                this.Error = "";
                oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oOrders);

                oPedido.CardCode = "Cliente01";
                oPedido.DocDate = DateTime.Today;
                oPedido.DocDueDate = DateTime.Today;
                oPedido.Comments = "Creado desde DI API";

                oPedido.Lines.ItemCode = "A00001";
                oPedido.Lines.Quantity = 5;
                oPedido.Lines.TaxCode = "IVA";

                oPedido.Lines.Add();

                oPedido.Lines.ItemCode = "A00002";
                oPedido.Lines.Quantity = 2;
                oPedido.Lines.TaxCode = "IVA";

                if (oPedido.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription() + " (" + this.oCom.GetLastErrorCode() + " )";
                }else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }

            }
            catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oPedido != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPedido);
                    oPedido = null;
                }
            }
        }

        public void AgregarLineaPedido(int DocEntry)
        {
            SAPbobsCOM.Documents oPedido = null;
            try
            {
                this.Error = "";
                oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oOrders);

                if (oPedido.GetByKey(DocEntry))
                {
                    if(oPedido.DocumentStatus!= SAPbobsCOM.BoStatus.bost_Close)
                    {
                        oPedido.Lines.Add();
                        oPedido.Lines.ItemCode = "A00003";
                        oPedido.Lines.Quantity = 1;
                        oPedido.Lines.TaxCode = "EXE";

                        if (oPedido.Update() != 0)
                        {
                            this.Error= this.oCom.GetLastErrorDescription() + " (" + this.oCom.GetLastErrorCode() + " )";
                        }
                    }else
                    {
                        this.Error = "Documento ya esta cerrado";
                    }
                }else
                {
                    this.Error = "Documento no existe";
                }

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oPedido != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPedido);
                    oPedido = null;
                }
            }
        }

        public void AgregarPedidoTipoServicio(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oPedido = null;
            try
            {
                oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oOrders);

                oPedido.CardCode = "Cliente01";
                oPedido.DocDate = DateTime.Today;
                oPedido.DocDueDate = DateTime.Today;
                oPedido.Comments = "Pedido de tipo servicio creado por DI API";

                oPedido.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;

                oPedido.Lines.ItemDescription = "Servicio de ejemplo";
                oPedido.Lines.AccountCode = "_SYS00000000001";
                oPedido.Lines.LineTotal = 300;

                if (oPedido.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription() + " (" + this.oCom.GetLastErrorCode() + ") ";
                }
                else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }


            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oPedido != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPedido);
                    oPedido = null;
                }
            }
        }


        public void CrearEntrega(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oEntrega = null;
            try
            {
                this.Error = "";
                oEntrega = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                oEntrega.CardCode = "Cliente01";
                oEntrega.DocDate = DateTime.Today;
                oEntrega.DocDueDate = DateTime.Today;
                oEntrega.Comments = "Entrega Creada desde DI API";
                oEntrega.UserFields.Fields.Item("U_Comentario").Value = "Creando entrega desde DI API en UDF";

                oEntrega.Lines.ItemCode = "A00001";
                oEntrega.Lines.Quantity = 2;
                oEntrega.Lines.TaxCode = "IVA";

                oEntrega.Lines.Add();

                oEntrega.Lines.ItemCode = "A00002";
                oEntrega.Lines.Quantity = 1;
                oEntrega.Lines.TaxCode = "IVA";

                if (oEntrega.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription() + " (" + this.oCom.GetLastErrorCode() + ") ";
                }else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }

            }
            catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oEntrega != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEntrega);
                    oEntrega = null;
                }
            }
        }


        public void CrearDevolucion(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oDevolucion = null;
            try
            {
                this.Error = "";
                oDevolucion = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oReturns);

                oDevolucion.CardCode = "Cliente01";
                oDevolucion.DocDate = DateTime.Today;
                oDevolucion.DocDueDate = DateTime.Today;

                oDevolucion.Lines.ItemCode = "A00001";
                oDevolucion.Lines.Quantity = 2;
                oDevolucion.Lines.TaxCode = "IVA";
                oDevolucion.Lines.UserFields.Fields.Item("U_Dscpton").Value="Valor de ejemplo";

                if (oDevolucion.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription() + " (" + this.oCom.GetLastErrorCode() + ") ";
                }
                else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }

            }
            catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oDevolucion != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDevolucion);
                    oDevolucion = null;
                }
            }
        }


        public void CrearSalida(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oSalida = null;
            try
            {
                this.Error = "";
                oSalida = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oInventoryGenExit);
                oSalida.DocDate = DateTime.Today;
                oSalida.DocDueDate = DateTime.Today;
                oSalida.GroupNumber =-1;
                

                oSalida.Lines.ItemCode = "B10000";
                oSalida.Lines.Quantity = 4;
                oSalida.Lines.CostingCode = "10001";
                oSalida.Lines.CostingCode2 = "20001";

                oSalida.Lines.BatchNumbers.SetCurrentLine(0);
                oSalida.Lines.BatchNumbers.Quantity = 4;
                oSalida.Lines.BatchNumbers.BatchNumber = "B1-00067";
                

                if (oSalida.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }


            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSalida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalida);
                   oSalida = null;
                }
            }
        }

        public void CrearFacturaConDocumentoBase(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oFactura = null;
            try
            {
                this.Error = "";
                oFactura = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oInvoices);
                oFactura.CardCode = "Cliente01";
                oFactura.DocDate = DateTime.Today;
                oFactura.DocDueDate = DateTime.Today;

                oFactura.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oOrders;
                oFactura.Lines.BaseEntry = 564;
                oFactura.Lines.BaseLine = 0;
                oFactura.Lines.TaxCode = "IVA";

                oFactura.Lines.Add();
                oFactura.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oOrders;
                oFactura.Lines.BaseEntry = 564;
                oFactura.Lines.BaseLine = 1;
                oFactura.Lines.TaxCode = "IVA";

                oFactura.Lines.Add();
                oFactura.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oOrders;
                oFactura.Lines.BaseEntry = 564;
                oFactura.Lines.BaseLine = 2;
                oFactura.Lines.TaxCode = "IVA";

                if (oFactura.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }

            }
            catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oFactura != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFactura);
                    oFactura = null;
                }
            }
        }

        public void CrearPedidoEnBaseABorrador(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oPedido = null;
            SAPbobsCOM.Documents Draft = null;
            try
            {
                this.Error = "";
                oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oOrders);
                oPedido.CardCode = "Cliente01";
                oPedido.DocDate = DateTime.Today;
                oPedido.DocDueDate = DateTime.Today;

                oPedido.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oDrafts;
                oPedido.Lines.BaseEntry = 81;
                oPedido.Lines.BaseLine = 0;
                oPedido.Lines.TaxCode = "IVA";

                if (oPedido.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }
                else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oPedido != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPedido);
                    oPedido = null;
                }
            }
        }


        public void CrearTransferencia(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.StockTransfer oTransferencia = null;
            try
            {
                this.Error = "";
                oTransferencia = (SAPbobsCOM.StockTransfer)this.oCom.GetBusinessObject(BoObjectTypes.oStockTransfer);
                oTransferencia.CardCode = "Cliente01";
                oTransferencia.DocDate = DateTime.Today;
                oTransferencia.TaxDate = DateTime.Today;

                oTransferencia.FromWarehouse = "01";
                oTransferencia.ToWarehouse = "02";

                oTransferencia.Lines.ItemCode = "A00001";
                oTransferencia.Lines.Quantity = 5;
                oTransferencia.Lines.FromWarehouseCode = "01";
                oTransferencia.Lines.WarehouseCode = "5";

                oTransferencia.Lines.Add();
                oTransferencia.Lines.ItemCode = "A00002";
                oTransferencia.Lines.Quantity = 5;


                if (oTransferencia.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }
                else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oTransferencia != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oTransferencia);
                    oTransferencia = null;
                }
            }
        }


        #endregion

    }
}
