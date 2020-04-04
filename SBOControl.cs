using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;

namespace Llacolen_SBOService
{
   sealed class SBOControl
   {

      //  public static SAPbobsCOM.Documents oOrder; // Order object
      //  public static SAPbobsCOM.Documents oInvoice; // Invoice Object
      public static SAPbobsCOM.Recordset oRecordSet; // A recordset object
      public static SAPbobsCOM.Company oCompany; // The company object
      //public static System.Data.SqlClient.SqlConnection BDBatchSQLServer; // Batch SQLServer database llacolen
      //public static MySqlConnection BDBatchMySQL;  // Batch MySQL database
      public ConexionMysql conexionMysql; 

      //Properties
      // Log Property
      private Logs.Logger FLog;
      public Logs.Logger oLog
      {
         get { return this.FLog; }
         set { this.FLog = value; }
      }


      //Methods
      public  void Doit(ref int nErr, ref string sErr )
      {
         if (oCompany == null)
            oCompany = new SAPbobsCOM.Company();

         if (!oCompany.Connected)
         {
            ConnectSBO(ref nErr, ref sErr);
            if (nErr != 0)
            {
                oLog.LogMsg(sErr, "F", "E");
                return;
            }
            else
            {
             oLog.LogMsg("Connected to SAP", "F", "D");
            }
         }

         //if (Llacolen_SBOService.Properties.Settings.Default.BatchMySQLType == "SQLServer")
         //{
         //    if (BDBatchSQLServer == null)
         //        BDBatchSQLServer = new System.Data.SqlClient.SqlConnection();

         //    if ((BDBatchSQLServer.State == System.Data.ConnectionState.Closed) | (BDBatchSQLServer.State == System.Data.ConnectionState.Broken))
         //    {
         //        ConnectBDBatchSQLServer(ref nErr, ref sErr);
         //        if (nErr != 0)
         //        {
         //            oLog.LogMsg(sErr, "F", "E");
         //            return;
         //        }
         //    }
         //}
         //else
         //{
             //if (BDBatchMySQL == null)
             //    BDBatchMySQL = new MySqlConnection();

             //if ((BDBatchMySQL.State == System.Data.ConnectionState.Closed) | (BDBatchMySQL.State == System.Data.ConnectionState.Broken))
             //{
             //    ConnectBDBatchMySQL(ref nErr, ref sErr);
             //    if (nErr != 0)
             //    {
             //        oLog.LogMsg(sErr, "F", "E");
             //        return;
             //    }
             //    else
             //    {
             //        oLog.LogMsg("Connected to MYSQL", "F", "D");
             //    }
             //}
         //}

         oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
         LlacolenDocs SincDocs  = new LlacolenDocs();
         SincDocs.oCompany      = oCompany;
        // SincDocs.oBDBatchMySQL = conexionMysql;
         SincDocs.oLog          = FLog;
         SincDocs.oRecordSet    = oRecordSet;
         SincDocs.InitVars();

         conexionMysql = new ConexionMysql();
         conexionMysql.oLog = FLog;

         oLog.LogMsg("------------------PROCESO 1-----------------------","F","E");
         oLog.LogMsg("Bloqueo/desbloqueo de socios en base a pagos, iniciado.", "F", "E");
 
         SincDocs.Sincronizar_PagosSAP_en_MySQL();

         oLog.LogMsg("------------------PROCESO 2-----------------------", "F", "E");
         oLog.LogMsg("Actualizar tickets e ingresar socios nuevos a punto, iniciado.", "F", "E");

         SincDocs.ActualizarSocios_MySQL();

         oLog.LogMsg("------------------PROCESO 3-----------------------", "F", "E");   
         oLog.LogMsg("Actualizar cupos desde punto, iniciado.", "F", "E");
 
         SincDocs.ActualizarCuposBP_SBO();
         
         oLog.LogMsg("Finaliza actualización.", "F", "E");
      }

      private void ConnectSBO(ref int nErr, ref string sErr)
      {
         sErr = "";
         nErr = 0;

         Llacolen_SBOService.Properties.Settings.Default.Reload();
         // Set connection properties
         if (Llacolen_SBOService.Properties.Settings.Default.SQLType == 2005)
         {
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
         }
         else if (Llacolen_SBOService.Properties.Settings.Default.SQLType == 2008)
         {
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
         }
         else if (Llacolen_SBOService.Properties.Settings.Default.SQLType == 2012)
         {
             oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
         }
         else if (Llacolen_SBOService.Properties.Settings.Default.SQLType == 2014)
         {
             oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
         }
         else if (Llacolen_SBOService.Properties.Settings.Default.SQLType == 2016)
         {
             oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
         }

         oCompany.Server = Llacolen_SBOService.Properties.Settings.Default.Server;
         oCompany.UseTrusted = false;
         oCompany.DbUserName = Llacolen_SBOService.Properties.Settings.Default.DBUserName;
         oCompany.DbPassword = Llacolen_SBOService.Properties.Settings.Default.DBPassword;

         oCompany.CompanyDB = Llacolen_SBOService.Properties.Settings.Default.CompanyDB;
         oCompany.UserName = Llacolen_SBOService.Properties.Settings.Default.UserName;
         oCompany.Password = Llacolen_SBOService.Properties.Settings.Default.Password;

         //Try to connect
         nErr = oCompany.Connect();

         if (nErr != 0) // if the connection failed
            oCompany.GetLastError(out nErr, out sErr);
         else
            sErr = "OK";

         oLog.LogMsg("Conexión SBO " + sErr, "F", "I");
      }

      //private void ConnectBDBatchSQLServer(ref int nErr, ref string sErr)
      //{
      //   sErr = "";
      //   nErr = 0;

      //   BDBatchSQLServer.ConnectionString = "Data Source=" + Llacolen_SBOService.Properties.Settings.Default.BatchServer +
      //                                       ";Initial Catalog=" + Llacolen_SBOService.Properties.Settings.Default.BatchDB +
      //                                       ";User ID=" + Llacolen_SBOService.Properties.Settings.Default.BatchUser +
      //                                       ";Password=" + Llacolen_SBOService.Properties.Settings.Default.BatchPassword +
      //                                       ";MultipleActiveResultSets=True;";
      //   try
      //   {
      //      BDBatchSQLServer.Open();
      //      sErr = "Conexión BD Batch SQLServer - OK";
      //      oLog.LogMsg(sErr, "F", "I");
      //   }
      //   catch (Exception e)
      //   {
      //      sErr = "BD Batch SQLServer " + e.Message;
      //      nErr = -1;
      //      oLog.LogMsg(sErr, "F", "E");
      //   }
      //}

      //private void ConnectBDBatchMySQL(ref int nErr, ref string sErr)
      //{
      //   sErr = "";
      //   nErr = 0;

      //   //
      //   BDBatchMySQL.ConnectionString = "Server   =" + Llacolen_SBOService.Properties.Settings.Default.BatchServer +
      //                                   ";port     =" + Llacolen_SBOService.Properties.Settings.Default.BatchMySQLPort +
      //                                   ";Database =" + Llacolen_SBOService.Properties.Settings.Default.BatchDB +
      //                                   ";Uid      =" + Llacolen_SBOService.Properties.Settings.Default.BatchUser +
      //                                   ";Password =" + Llacolen_SBOService.Properties.Settings.Default.BatchPassword +
      //                                   "";
      //   try
      //   {
      //      BDBatchMySQL.Open();
      //      sErr = "Conexión BD Batch MySQL - OK";
      //      oLog.LogMsg(sErr, "F", "I");
      //   }
      //   catch (MySqlException e)
      //   {
      //      sErr = "BD Batch MySQL " + e.Message;
      //      nErr = -1;
      //      oLog.LogMsg(sErr, "F", "E");
      //   }
      //}

   }
}
