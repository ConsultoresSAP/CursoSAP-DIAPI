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


    }
}
