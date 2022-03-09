using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using System.Xml;

namespace CursoSDK_DIAPI
{
    class ValoresValidos
    {
        public string Code { get; set; }
        public string Desc { get; set; }
    }

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
                    } else
                    {
                        this.CName = this.oCom.CompanyName;
                    }
                }


            } catch (Exception e)
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

            } catch (Exception e)
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

            } catch (Exception e)
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
                } else
                {
                    this.Error = "SN no existe";
                }


            } catch (Exception e)
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



        public void AddContactoSN(string CardCode, string Name)
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
                    } else
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

                } else
                {
                    this.Error = "SN no existe";
                }

            } catch (Exception e)
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


        public void EditContacto(string CardCode, int line)
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
            catch (Exception e)
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
            catch (Exception e)
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

            } catch (Exception e)
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

        public void ActualizarListaDePrecios(int Lista, int ListaBase,string ItemCode)
        {
            SAPbobsCOM.Items Item = null;
            try
            {
                this.Error = "";
                Item = (SAPbobsCOM.Items)this.oCom.GetBusinessObject(BoObjectTypes.oItems);

                double precio = 400;
                double Factor = 2;

                if (Item.GetByKey(ItemCode))
                {
                    Item.PriceList.SetCurrentLine(Lista);
                    Item.PriceList.BasePriceList = ListaBase;
                    Item.PriceList.Price = precio;
                    Item.PriceList.Factor = Factor;

                    if (Item.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }else
                {
                    this.Error = "Articulo no existe";
                }

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {

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
                } else
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

        public void AgregarLineaPedido(int DocEntry)
        {
            SAPbobsCOM.Documents oPedido = null;
            try
            {
                this.Error = "";
                oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oOrders);

                if (oPedido.GetByKey(DocEntry))
                {
                    if (oPedido.DocumentStatus != SAPbobsCOM.BoStatus.bost_Close)
                    {
                        oPedido.Lines.Add();
                        oPedido.Lines.ItemCode = "A00003";
                        oPedido.Lines.Quantity = 1;
                        oPedido.Lines.TaxCode = "EXE";

                        if (oPedido.Update() != 0)
                        {
                            this.Error = this.oCom.GetLastErrorDescription() + " (" + this.oCom.GetLastErrorCode() + " )";
                        }
                    } else
                    {
                        this.Error = "Documento ya esta cerrado";
                    }
                } else
                {
                    this.Error = "Documento no existe";
                }

            } catch (Exception e)
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


            } catch (Exception e)
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
                } else
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
                oDevolucion.Lines.UserFields.Fields.Item("U_Dscpton").Value = "Valor de ejemplo";

                if (oDevolucion.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription() + " (" + this.oCom.GetLastErrorCode() + ") ";
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
                oSalida.GroupNumber = -1;


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
                } else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }


            } catch (Exception e)
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
                } else
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
                if (oFactura != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFactura);
                    oFactura = null;
                }
            }
        }

        public void CrearFacturaConDocumentoBase(string DocNumPedido,out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oFactura = null;
            SAPbobsCOM.Recordset oRecord = null;
            try
            {
                this.Error = "";
                oFactura = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oInvoices);
                oRecord = (SAPbobsCOM.Recordset)this.oCom.GetBusinessObject(BoObjectTypes.BoRecordset);

                if (this.oCom.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    oRecord.DoQuery("CALL PedidoBase  ('" + DocNumPedido+"')");
                }
                else
                {
                    oRecord.DoQuery("EXEC PedidoBase " + DocNumPedido);
                }
                

                if (oRecord.RecordCount > 0)
                {
                    oRecord.MoveFirst();
                    oFactura.CardCode = oRecord.Fields.Item("SN").Value.ToString();
                    oFactura.DocDate = DateTime.Today;
                    oFactura.DocDueDate = DateTime.Today;

                    for(int i = 0; i < oRecord.RecordCount; i++)
                    {
                        if (i != 0)
                        {
                            oFactura.Lines.Add();
                        }
                        oFactura.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oOrders;
                        oFactura.Lines.BaseEntry = Int32.Parse(oRecord.Fields.Item("N. Interno").Value.ToString());
                        oFactura.Lines.BaseLine = Int32.Parse(oRecord.Fields.Item("Linea").Value.ToString());
                        oFactura.Lines.TaxCode = oRecord.Fields.Item("Impuesto").Value.ToString();
                        oRecord.MoveNext();
                    }

                    if (oFactura.Add() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                    else
                    {
                        DocEntry = this.oCom.GetNewObjectKey();
                    }
                }
                else
                {
                    this.Error = "Pedido no encontrado";
                }

                

            }
            catch (Exception e)
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
                if (oRecord != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord);
                    oRecord = null;
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
                this.oCom.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                this.oCom.XMLAsString = false;
                string rutaXml = @"C:\Users\User-02\Desktop\CursoSAP DI API\CursoSDK-DIAPI\XmlBorrador.xml";
                XmlDocument XmlBorrador = new XmlDocument();
                XmlNode NodeDocObjectCode;
                XmlNode NodeDocNum;
                XmlNode Node;
                Draft = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oDrafts);

                if (Draft.GetByKey(88))
                {
                    Draft.SaveXML(ref rutaXml);

                    XmlBorrador.Load(rutaXml);
                    XmlBorrador.SelectSingleNode("BOM/BO/AdmInfo/Object").InnerText = "17";
                    NodeDocNum = XmlBorrador.SelectSingleNode("BOM/BO/Documents/row/DocNum");
                    XmlBorrador.SelectSingleNode("BOM/BO/Documents/row").RemoveChild(NodeDocNum);

                    NodeDocObjectCode = XmlBorrador.SelectSingleNode("BOM/BO/Documents/row/DocObjectCode");
                    XmlBorrador.SelectSingleNode("BOM/BO/Documents/row").RemoveChild(NodeDocObjectCode);

                    Node = XmlBorrador.SelectSingleNode("BOM/BO/Documents/row/ReqType");
                    XmlBorrador.SelectSingleNode("BOM/BO/Documents/row").RemoveChild(Node);
                    Node = XmlBorrador.SelectSingleNode("BOM/BO/Documents/row/Revision");
                    XmlBorrador.SelectSingleNode("BOM/BO/Documents/row").RemoveChild(Node);
                    Node = XmlBorrador.SelectSingleNode("BOM/BO/Documents/row/IssuingReason");
                    XmlBorrador.SelectSingleNode("BOM/BO/Documents/row").RemoveChild(Node);

                    for (int i = 0; i < XmlBorrador.SelectNodes("BOM/BO/Document_Lines/row").Count; i++)
                    {
                        var NodoActual = XmlBorrador.SelectNodes("BOM/BO/Document_Lines/row")[i];
                        Node = NodoActual.SelectSingleNode("EnableReturnCost");
                        XmlBorrador.SelectNodes("BOM/BO/Document_Lines/row")[i].RemoveChild(Node);
                        Node = NodoActual.SelectSingleNode("ReturnCost");
                        XmlBorrador.SelectNodes("BOM/BO/Document_Lines/row")[i].RemoveChild(Node);
                        Node = NodoActual.SelectSingleNode("LineVendor");
                        XmlBorrador.SelectNodes("BOM/BO/Document_Lines/row")[i].RemoveChild(Node);
                    }


                    XmlBorrador.Save(rutaXml);

                    oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObjectFromXML(rutaXml, 0);

                    if (oPedido.Add() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                    else
                    {
                        DocEntry = this.oCom.GetNewObjectKey();
                        Draft.Remove();
                    }
                } else
                {
                    this.Error = "Borrador no existe";
                }


            }
            catch (System.Runtime.InteropServices.COMException e)
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
                if (Draft != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Draft);
                    Draft = null;
                }
            }
        }





        #endregion

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

        public void CrearPago(int DocEntryFact, out string DocEntryPago)
        {
            DocEntryPago = "";
            SAPbobsCOM.Payments oPago = null;
            SAPbobsCOM.Documents oFactura = null;
            try
            {
                this.Error = "";
                oPago = (SAPbobsCOM.Payments)this.oCom.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                oFactura = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oInvoices);

                if (oFactura.GetByKey(DocEntryFact))
                {
                    oPago.CardCode = oFactura.CardCode;
                    oPago.DocDate = DateTime.Today;
                    oPago.DueDate = DateTime.Today;

                    //Medios de Pago

                    //Efectivo
                    //oPago.CashSum = oFactura.DocTotal;
                    oPago.CashSum = 1000;

                    //Transferencias
                    oPago.TransferAccount = "_SYS00000000001";
                    oPago.TransferDate = DateTime.Today;
                    oPago.TransferReference = "89452315";
                    oPago.TransferSum = 1000;

                    //Cheques
                    oPago.Checks.DueDate = DateTime.Today;
                    oPago.Checks.CheckSum = 1000;
                    oPago.Checks.CountryCode = "GT";
                    oPago.Checks.BankCode = "BBANK";
                    oPago.Checks.AccounttNum = "89451265785";
                    oPago.Checks.CheckNumber = 894523;
                    oPago.Checks.Trnsfrable = SAPbobsCOM.BoYesNoEnum.tYES;


                    //TC
                    oPago.CreditCards.CreditCard = 2;
                    oPago.CreditCards.CreditCardNumber = "1234567891232541";
                    oPago.CreditCards.CardValidUntil = DateTime.Parse("10/10/26");
                    oPago.CreditCards.VoucherNum = "4523152";
                    oPago.CreditCards.CreditSum = 1000;

                    oPago.Invoices.DocEntry = DocEntryFact;
                    oPago.Invoices.SumApplied = 4000;

                    if (oPago.Add() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                    else
                    {
                        DocEntryPago = this.oCom.GetNewObjectKey();
                    }

                }
                else
                {
                    this.Error = "Factura no existe";
                }

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {

            }
        }

        public void CrearPago(string DocNumFact, out string DocEntryPago)
        {
            DocEntryPago = "";
            SAPbobsCOM.Payments oPago = null;
            SAPbobsCOM.Recordset oRecord = null;
            try
            {
                this.Error = "";
                oPago = (SAPbobsCOM.Payments)this.oCom.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                oRecord = (SAPbobsCOM.Recordset)this.oCom.GetBusinessObject(BoObjectTypes.BoRecordset);

                oRecord.DoQuery("SELECT (T0.DocTotal-T0.PaidToDate) AS 'Pendiente',T0.[CardCode],T0.[DocEntry] " +
                                "FROM [OINV] T0 WHERE T0.[DocNum]= " + DocNumFact);

                if (oRecord.RecordCount > 0)
                {
                    oPago.CardCode = oRecord.Fields.Item("CardCode").Value.ToString();
                    oPago.DocDate = DateTime.Today;
                    oPago.DueDate = DateTime.Today;

                    oPago.CashSum = Double.Parse(oRecord.Fields.Item("Pendiente").Value.ToString());



                    oPago.Invoices.DocEntry = Int32.Parse(oRecord.Fields.Item("DocEntry").Value.ToString());
                    oPago.Invoices.SumApplied = Double.Parse(oRecord.Fields.Item("Pendiente").Value.ToString());

                    if (oPago.Add() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                    else
                    {
                        DocEntryPago = this.oCom.GetNewObjectKey();
                    }

                }
                else
                {
                    this.Error = "Factura no existe";
                }

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oPago != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPago);
                    oPago = null;
                }
            }

        }

        public void Record(int DocEntryFact, out string Datos)
        {
            Datos = "";
            SAPbobsCOM.Recordset oRecord = null;
            try
            {
                this.Error = "";
                oRecord = (SAPbobsCOM.Recordset)this.oCom.GetBusinessObject(BoObjectTypes.BoRecordset);
                string Consulta = "";
                if (this.oCom.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Consulta = "SELECT * FROM \"OINV\" T0 WHERE T0.\"DocEntry\"= " + DocEntryFact.ToString();
                }
                else
                {
                    Consulta = "SELECT * FROM [dbo].[OINV] T0 WHERE T0.[DocEntry]= " + DocEntryFact.ToString();
                }


                oRecord.DoQuery(Consulta);

                if (oRecord.RecordCount > 0)
                {
                    Datos = "Cliente: " + oRecord.Fields.Item("CardCode").Value.ToString() + "-" + oRecord.Fields.Item("CardName").Value.ToString() +
                        ", Total Factuta: " + oRecord.Fields.Item("DocTotal").Value.ToString();
                }
                else
                {
                    this.Error = "No hay datos";
                }

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oRecord != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord);
                    oRecord = null;
                }
            }
        }


        public void CrearTabla(string Nombre, string Desc, SAPbobsCOM.BoUTBTableType Type)
        {
            SAPbobsCOM.IUserTablesMD oTabla = null;
            try
            {
                this.Error = "";
                oTabla = (SAPbobsCOM.IUserTablesMD)this.oCom.GetBusinessObject(BoObjectTypes.oUserTables);

                if (!oTabla.GetByKey(Nombre))
                {
                    oTabla.TableName = Nombre;
                    oTabla.TableDescription = Desc;
                    oTabla.TableType = Type;

                    if (oTabla.Add() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }else
                {
                    this.Error = "Tabla ya existe";
                }

            }catch(System.Runtime.InteropServices.COMException e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oTabla != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oTabla);
                    oTabla = null;
                    GC.Collect();
                }
            }
        }


        public void CrearOActualizarUDF(string Tabla, string Code, string Desc, int Tam,
                            SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType,
                            SAPbobsCOM.BoYesNoEnum Obligatorio, string TablaEnlazada,
                            List<ValoresValidos> ValoresValidos,string ValorDefecto)
        {
            SAPbobsCOM.UserFieldsMD oUDF = null;
            SAPbobsCOM.Recordset oRecord = null;
            try
            {
                this.Error = "";
                oUDF = (SAPbobsCOM.UserFieldsMD)this.oCom.GetBusinessObject(BoObjectTypes.oUserFields);
                oRecord = (SAPbobsCOM.Recordset)this.oCom.GetBusinessObject(BoObjectTypes.BoRecordset);

                int Key;
                bool Existe = false;
                oRecord.DoQuery("SELECT T0.[FieldID] FROM [CUFD] T0 WHERE T0.[TableID]='" + Tabla + "' AND T0.[AliasID]='" + Code + "'");
                if (oRecord.RecordCount > 0)
                {
                    Key = Int32.Parse(oRecord.Fields.Item("FieldID").Value.ToString());
                    oUDF.GetByKey(Tabla, Key);
                    Existe = true;
                }

                if (oRecord != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord);
                    oRecord = null;
                }

                oUDF.TableName = Tabla;
                oUDF.Name = Code;
                oUDF.Description = Desc;
                oUDF.Size = Tam;
                oUDF.Type = Type;
                oUDF.SubType = SubType;
                oUDF.Mandatory = Obligatorio;

                if (TablaEnlazada != "")
                {
                    oUDF.LinkedTable = TablaEnlazada;
                }
                if (ValorDefecto != "")
                {
                    oUDF.DefaultValue = ValorDefecto;
                }

                if (ValoresValidos != null)
                {
                    for(int i = 0; i < ValoresValidos.Count; i++)
                    {
                        oUDF.ValidValues.Value = ValoresValidos[i].Code;
                        oUDF.ValidValues.Description = ValoresValidos[i].Desc;
                        oUDF.ValidValues.Add();
                    }
                }

                if (Existe)
                {
                    if (oUDF.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }else
                {
                    if (oUDF.Add() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }

            }catch(System.Runtime.InteropServices.COMException e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oUDF != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                    oUDF = null;
                    GC.Collect();
                }
            }
        }


        public void CrearUDO()
        {
            SAPbobsCOM.UserObjectsMD UDO = null;
            try
            {
                this.Error = "";
                UDO = (SAPbobsCOM.UserObjectsMD)this.oCom.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

                UDO.Code = "UDO_EJM";
                UDO.Name = "UDO de ejemplo";
                UDO.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document;
                UDO.TableName = "TABLAPADRE";

                UDO.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                UDO.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                UDO.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                UDO.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                UDO.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;

                //UDO.FindColumns.ColumnAlias = "U_NUMERO";
                //UDO.FindColumns.Add();

                UDO.ChildTables.TableName = "TablaHija";
                UDO.ChildTables.Add();

                if (UDO.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }

            }catch(System.Runtime.InteropServices.COMException e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (UDO != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(UDO);
                    UDO = null;
                }
            }
        }

        public void AgregarDatosUDO()
        {
            SAPbobsCOM.CompanyService OCompanyServices = null;
            SAPbobsCOM.GeneralService OGeneralServices = null;
            SAPbobsCOM.GeneralData OGeneralData = null;
            SAPbobsCOM.GeneralDataParams OGeneralDataParams = null;

            SAPbobsCOM.GeneralDataCollection Lineas = null;
            SAPbobsCOM.GeneralData Linea = null;
            try
            {
                this.Error = "";
                OCompanyServices = this.oCom.GetCompanyService();
                OGeneralServices = OCompanyServices.GetGeneralService("VISITVEN");

                OGeneralData = ((SAPbobsCOM.GeneralData)(OGeneralServices.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                OGeneralData.SetProperty("U_CodVendedor", "V0001");
                OGeneralData.SetProperty("U_Nombre", "Juan Perez");
                OGeneralData.SetProperty("U_Fecha", DateTime.Today);
                OGeneralData.SetProperty("U_Comentarios", "Visitas de hoy");

                Lineas = OGeneralData.Child("VISITASVENDEDOR1");

                Linea = Lineas.Add();
                Linea.SetProperty("U_CodCliente", "Cliente01");
                Linea.SetProperty("U_Nombre", "Cliente de prueba numero");
                Linea.SetProperty("U_Usunto", "Cosas Varias");//Asunto
                Linea.SetProperty("U_Preoridad", "Alta");//Prioridad

                OGeneralDataParams = OGeneralServices.Add(OGeneralData);

            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (OCompanyServices != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OCompanyServices);
                    OCompanyServices = null;
                }
                if (OGeneralServices != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OGeneralServices);
                    OGeneralServices = null;
                }
                if (OGeneralData != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OGeneralData);
                    OGeneralData = null;
                }
                if (OGeneralDataParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OGeneralDataParams);
                    OGeneralDataParams = null;
                }
                if (Lineas != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Lineas);
                    Lineas = null;
                }
                if(Linea != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Linea);
                    Linea = null;
                }
            }
        }

        public void EditarDatosUDO(int DocEntry)
        {
            SAPbobsCOM.CompanyService OCompanyServices = null;
            SAPbobsCOM.GeneralService OGeneralServices = null;
            SAPbobsCOM.GeneralData OGeneralData = null;
            SAPbobsCOM.GeneralDataParams OGeneralDataParams = null;

            SAPbobsCOM.GeneralDataCollection Lineas = null;
            SAPbobsCOM.GeneralData Linea = null;
            try
            {
                this.Error = "";
                OCompanyServices = this.oCom.GetCompanyService();
                OGeneralServices = OCompanyServices.GetGeneralService("VISITVEN");

                OGeneralDataParams = ((SAPbobsCOM.GeneralDataParams)(OGeneralServices.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));
                OGeneralDataParams.SetProperty("DocEntry", DocEntry);

                OGeneralData = OGeneralServices.GetByParams(OGeneralDataParams);

                OGeneralData.SetProperty("U_Comentarios", "Visitas de hoy, con nuevo cliente");

                Lineas = OGeneralData.Child("VISITASVENDEDOR1");

                //Agregar Linea
                Linea = Lineas.Add();
                Linea.SetProperty("U_CodCliente", "CL03");
                Linea.SetProperty("U_Nombre", "Cliente TEST");
                Linea.SetProperty("U_Usunto", "Cosas Varias");//Asunto
                Linea.SetProperty("U_Preoridad", "Baja");//Prioridad

                //Editar linea
                //Linea = Lineas.Item(0);
                //Linea.SetProperty("U_Preoridad", "Baja");//Prioridad

                OGeneralServices.Update(OGeneralData);

            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (OCompanyServices != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OCompanyServices);
                    OCompanyServices = null;
                }
                if (OGeneralServices != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OGeneralServices);
                    OGeneralServices = null;
                }
                if (OGeneralData != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OGeneralData);
                    OGeneralData = null;
                }
                if (OGeneralDataParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(OGeneralDataParams);
                    OGeneralDataParams = null;
                }
                if (Lineas != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Lineas);
                    Lineas = null;
                }
                if (Linea != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Linea);
                    Linea = null;
                }
            }
        }


        public void EjemploTransaction()
        {
            SAPbobsCOM.Documents Factura = null;
            SAPbobsCOM.Payments Pago = null;
            try
            {
                this.Error = "";
                string DocEntry = "";
                Factura = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oInvoices);
                Pago = (SAPbobsCOM.Payments)this.oCom.GetBusinessObject(BoObjectTypes.oIncomingPayments);

                if (!this.oCom.InTransaction)
                {
                    this.oCom.StartTransaction();
                }

                if (this.oCom.InTransaction)
                {
                    Factura.CardCode = "Cliente01";
                    Factura.DocDate = DateTime.Today;
                    Factura.DocDueDate = DateTime.Today;
                    Factura.Comments = "Factura creada en proceso de Transaction 2";

                    Factura.Lines.ItemCode = "A00001";
                    Factura.Lines.Quantity = 2;
                    Factura.Lines.TaxCode = "IVA";

                    if (Factura.Add() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                        if (this.oCom.InTransaction)
                        {
                            this.oCom.EndTransaction(BoWfTransOpt.wf_RollBack);
                        }
                    }else
                    {
                        DocEntry = this.oCom.GetNewObjectKey();
                        if (Factura.GetByKey(Int32.Parse(DocEntry)))
                        {
                            Pago.CardCode = Factura.CardCode;
                            Pago.DocDate = DateTime.Today;

                            Pago.CashSum = Factura.DocTotal;

                            Pago.Invoices.DocEntry = Factura.DocEntry;
                            Pago.Invoices.TotalDiscount = 0;
                            Pago.Invoices.SumApplied = Factura.DocTotal;

                            if (Pago.Add() != 0)
                            {
                                this.Error = this.oCom.GetLastErrorDescription();
                                if (this.oCom.InTransaction)
                                {
                                    this.oCom.EndTransaction(BoWfTransOpt.wf_RollBack);
                                }
                            }else
                            {
                                if (this.oCom.InTransaction)
                                {
                                    this.oCom.EndTransaction(BoWfTransOpt.wf_Commit);
                                }
                            }
                        }
                        

                    }
                }
                

            }catch(Exception e)
            {
                this.Error = e.Message;
                if (this.oCom.InTransaction)
                {
                    this.oCom.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }finally
            {
                if (this.oCom != null)
                {
                    if (this.oCom.InTransaction)
                    {
                        this.oCom.EndTransaction(BoWfTransOpt.wf_RollBack);
                    }
                }
                if (Factura != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Factura);
                    Factura = null;
                }
                if (Pago != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Pago);
                    Pago = null;
                }
            }
        }


    }
}
