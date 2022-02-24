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




    }
}
