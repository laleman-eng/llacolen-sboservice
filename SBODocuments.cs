using System;
using MySql.Data.MySqlClient;


namespace Llacolen_SBOService
{
    public class LlacolenDocs
    {
        // Company Property
        private SAPbobsCOM.Company FCompany;
        public ConexionMysql conexionMysql;

        public SAPbobsCOM.Company oCompany
        {
            get { return this.FCompany; }
            set { this.FCompany = value; }
        }

        // Recordset Property
        private SAPbobsCOM.Recordset FRecordSet;
        public SAPbobsCOM.Recordset oRecordSet
        {
            get { return this.FRecordSet; }
            set { this.FRecordSet = value; }
        }

        // BDBatchSQLServer Property
        //private System.Data.SqlClient.SqlConnection FBDBatchSQLServer;
        //public System.Data.SqlClient.SqlConnection oBDBatchSQLServer
        //{
        //    get { return this.FBDBatchSQLServer; }
        //    set { this.FBDBatchSQLServer = value; }
        //}

        // BDBatchMySQL Property
        //private ConexionMysql FBDBatchMySQL;

        //public ConexionMysql oBDBatchMySQL
        //{
        //    get { return this.FBDBatchMySQL; }
        //    set { this.FBDBatchMySQL = value; }
        //}

        // Log Property
        private Logs.Logger FLog;
        public Logs.Logger oLog
        {
            get { return this.FLog; }
            set { this.FLog = value; }
        }

        // Var
        private SAPbobsCOM.BusinessPartners oBP;

        public void InitVars()
        {
            if (oCompany != null)
                oBP = (SAPbobsCOM.BusinessPartners)FCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            else
                new Exception("Compañia SBO, no encontrada.");
        }

        //private Boolean GetIntField(ref SqlDataReader Reader, String Name, ref int value, String sError)
        //{
        //    int i = Reader.GetOrdinal("KeyId");
        //    if (Reader.IsDBNull(i))
        //    {
        //        if (sError != null)
        //        {
        //            oLog.LogMsg(sError, "F", "D");
        //        }
        //        return false;
        //    }
        //    else
        //    {
        //        value = Reader.GetInt32(i);
        //        return true;
        //    }
        //}

        //private Boolean GetStringField(ref SqlDataReader Reader, String Name, ref String value, String sError)
        //{
        //    int i = Reader.GetOrdinal("KeyId");
        //    if (Reader.IsDBNull(i))
        //    {
        //        if (sError != null)
        //        {
        //            oLog.LogMsg(sError, "F", "D");
        //        }
        //        return false;
        //    }
        //    else
        //    {
        //        value = Reader.GetString(i);
        //        return true;
        //    }
        //}

        //private Boolean GetDateTimeField(ref SqlDataReader Reader, String Name, ref DateTime value, String sError)
        //{
        //    int i = Reader.GetOrdinal("KeyId");
        //    if (Reader.IsDBNull(i))
        //    {
        //        if (sError != null)
        //        {
        //            oLog.LogMsg(sError, "F", "D");
        //        }
        //        return false;
        //    }
        //    else
        //    {
        //        value = Reader.GetDateTime(i);
        //        return true;
        //    }
        //}

        //private Boolean GetDoubleField(ref SqlDataReader Reader, String Name, ref Double value, String sError)
        //{
        //    int i = Reader.GetOrdinal("KeyId");
        //    if (Reader.IsDBNull(i))
        //    {
        //        if (sError != null)
        //        {
        //            oLog.LogMsg(sError, "F", "D");
        //        }
        //        return false;
        //    }
        //    else
        //    {
        //        value = Reader.GetDouble(i);
        //        return true;
        //    }
        //}

        private string getLink()
        {
            string stringConnection = "Server    =" + Llacolen_SBOService.Properties.Settings.Default.BatchServer +
                                           ";port     =" + Llacolen_SBOService.Properties.Settings.Default.BatchMySQLPort +
                                           ";Database =" + Llacolen_SBOService.Properties.Settings.Default.BatchDB +
                                           ";Uid      =" + Llacolen_SBOService.Properties.Settings.Default.BatchUser +
                                           ";Password =" + Llacolen_SBOService.Properties.Settings.Default.BatchPassword + "";
            return stringConnection;
        }

        public void ActualizarCuposBP_SBO()
        {
            String oSql;
            //MySqlDataReader Reader;
            MySqlCommand Qry;
            String sErr;
            Int32 nErr;

            try
            {
                conexionMysql = new ConexionMysql();
                conexionMysql.oLog = FLog;

                oSql = "SELECT ACCESO_RUT, ACCESO_AUTORIZACION, ACCESO_CUPOS_VISITAS FROM acceso_llacolen";
                try
                {
                    if (conexionMysql.AbrirConexion() == true)
                    {
                        Qry = new MySqlCommand(oSql, conexionMysql.conexionMysql);

                        MySqlDataReader Reader = Qry.ExecuteReader();
                        
                        while (Reader.Read())
                        {
                            int cupos = Reader.GetInt32(2);
                            oSql = "Select CardCode " +
                                   "  From OCRD " +
                                   " Where (GroupCode = 100 or GroupCode = 102) " +
                                   "   and substring(LicTradNum,1,len(LicTradNum)-2)= '" + Reader.GetString(0) + "' " +
                                   "   and U_Tickets <> " + Reader.GetInt32(2).ToString();
                            FRecordSet.DoQuery(oSql);

                            if (!FRecordSet.EoF)
                            {
                                if (oBP.GetByKey((System.String)FRecordSet.Fields.Item("CardCode").Value))
                                {
                                    oBP.UserFields.Fields.Item("U_Tickets").Value = cupos;
                                    nErr = oBP.Update();
                                    if (nErr != 0)
                                    {
                                        sErr = oCompany.GetLastErrorDescription();
                                        oLog.LogMsg("Error al actualizar cupos desde MySQL a SAP, socio: " + (System.String)FRecordSet.Fields.Item("CardCode").Value + " :" + sErr, "A", "D");
                                    }
                                    else
                                    {
                                        oLog.LogMsg("Actualización de CUPOS en SBO, socio: " + (System.String)FRecordSet.Fields.Item("CardCode").Value, "A", "D");
                                    }
                                }
                                FRecordSet.MoveNext();
                            }
                        }
                        Reader.Close();
                    conexionMysql.CerrarConexion();
                    }
                }
                catch (MySqlException ex)
                {
                    oLog.LogMsg("MYSQLException " + ex.Message, "A", "D");
                    conexionMysql.CerrarConexion();
                    oLog.LogMsg("Cerrando Conexion EXCEPTION", "F", "E");
                }
            }
            catch (Exception e)
            {
                oLog.LogMsg("Error al actualizar cupos en SBO, problema con MySQL: " + e.Message, "A", "E");
            }
        }

        public void ActualizarSocios_MySQL()
        {
            String oSql;
            MySqlCommand QryCmd;
            //MySqlDataReader Reader;
            String Autorizacion;
            String Rut;
            Int32 TipoSocio;
            Int32 Tickets;

            try
            {
                conexionMysql = new ConexionMysql();
                conexionMysql.oLog = FLog;

                oSql = "Select CardCode, FrozenFor, isnull(U_Tickets,0) U_Tickets, LicTradNum, U_Ausente, GroupCode " +
                       "  from OCRD " +
                       " where CardType = 'C' and GroupCode <> '108'";  // 108 - grupo accionistas

                FRecordSet.DoQuery(oSql);
                while (!FRecordSet.EoF)
                {
                    Autorizacion = "Y";
                    if ((System.String)FRecordSet.Fields.Item("U_Ausente").Value == "2")  // U_Ausente = 2 => Si, 1 => No
                        Autorizacion = "N";
                    if ((System.String)FRecordSet.Fields.Item("FrozenFor").Value == "Y")
                        Autorizacion = "N";

                    Rut = (System.String)FRecordSet.Fields.Item("LicTradNum").Value;
                    Rut = Rut.Substring(0, Rut.Length - 2);
                    TipoSocio = (System.Int32)FRecordSet.Fields.Item("GroupCode").Value;
                    Tickets = (System.Int32)FRecordSet.Fields.Item("U_Tickets").Value;

                    try
                    {
                        if (conexionMysql.AbrirConexion() == true)
                        {
                            oSql = "SELECT ACCESO_AUTORIZACION, ACCESO_CUPOS_VISITAS FROM acceso_llacolen " +
                                   " WHERE ACCESO_RUT = '" + Rut + "'";
                            QryCmd = new MySqlCommand(oSql, conexionMysql.conexionMysql);
                            MySqlDataReader Reader = QryCmd.ExecuteReader();
                            if (Reader.HasRows)
                            {
                                while (Reader.Read())
                                {
                                    if (Autorizacion != Reader.GetString(0).Trim())
                                    {
                                        oSql = "UPDATE acceso_llacolen " +
                                               "   SET ACCESO_AUTORIZACION  = '" + Autorizacion + "'" +
                                               " WHERE ACCESO_RUT = '" + Rut + "'";
                                        MySqlConnection conn = new MySqlConnection(getLink());
                                        conn.Open();
                                        MySqlCommand QryCmd2 = new MySqlCommand(oSql, conn);
                                        if (QryCmd2.ExecuteNonQuery() > 0) 
                                            oLog.LogMsg("Actualización de AUTORIZACION sobre MySQL: " + Rut, "A", "D");
                                        conn.Close();
                                    }
                                    if ((Tickets > 0) && (Reader.GetInt32(1) == 0))
                                    {
                                        oSql = "UPDATE acceso_llacolen " +
                                               "   SET ACCESO_CUPOS_VISITAS =  " + Tickets.ToString() +
                                               " WHERE ACCESO_RUT = '" + Rut + "'";
                                        MySqlConnection conn = new MySqlConnection(getLink());
                                        conn.Open();
                                        MySqlCommand QryCmd2 = new MySqlCommand(oSql, conn);
                                        if (QryCmd2.ExecuteNonQuery() > 0)
                                            oLog.LogMsg("Actualización de TICKETS sobre MySQL: " + Rut, "A", "D");
                                        conn.Close();
                                    }
                                }
                                Reader.Close();
                            }
                            else
                            {
                                //SON CODIGOS CORRESPONDIENTES AL SOCIO TITULAR 100 Y CONYUJE 102, LLEVAN CUPOS
                                if ((TipoSocio == 100) || (TipoSocio == 102))
                                {
                                    oSql = "INSERT INTO acceso_llacolen " +
                                            "   ( ACCESO_RUT, ACCESO_AUTORIZACION, ACCESO_CUPOS_VISITAS) " +
                                            " VALUES (" +
                                            "'" + Rut + "'," +
                                            "'" + Autorizacion + "'," +
                                                Tickets.ToString() +
                                            ")";
                                }
                                else
                                {
                                    oSql = "INSERT INTO acceso_llacolen " +
                                            "   ( ACCESO_RUT, ACCESO_AUTORIZACION, ACCESO_CUPOS_VISITAS) " +
                                            " VALUES (" +
                                            "'" + Rut + "'," +
                                            "'" + Autorizacion + "'," +
                                                  "0" +
                                            ")";
                                }

                                MySqlConnection conn = new MySqlConnection(getLink());
                                conn.Open();
                                MySqlCommand QryCmd2 = new MySqlCommand(oSql, conn);
                                QryCmd2.ExecuteNonQuery();
                                oLog.LogMsg("Nuevo socio Ingresado a MySQL: " + (System.String)FRecordSet.Fields.Item("CardCode").Value, "A", "D");
                                conn.Close();
                            }
                            FRecordSet.MoveNext();
                        conexionMysql.CerrarConexion();
                        }
                    }
                    catch (MySqlException ex)
                    {
                        oLog.LogMsg("MYSQLException " + ex.Message, "A", "D");
                        conexionMysql.CerrarConexion();
                        oLog.LogMsg("Cerrando Conexion.", "F", "E");
                    }
                }  
            }
            catch (Exception e)
            {
                oLog.LogMsg("Error al actualizar socios en MySql : " + e.Message, "A", "E");
            }
        }

        public void Sincronizar_PagosSAP_en_MySQL()
        {
            String oSql;
            MySqlCommand QryCmd;
            //MySqlDataReader Reader;
            String Autorizacion;
            String Rut;
            String RutAux;
            String Tickets;
            String sErr;
            Int32 nErr;

            try
            {
                    conexionMysql = new ConexionMysql();
                    conexionMysql.oLog = FLog;
                    // Solo Deuda vencida => Bloquear
                    oSql = "Select DISTINCT spr.cardcode SocPrin, s.CardCode, s.LicTradNum, s.GroupCode, spr.saldo, s.frozenFor, s.U_Tickets " +
                           "  from OCRD s, " +
                           "       (select i.CardCode, i.DocTotal-i.PaidToDate as Saldo, i.DocDueDate, " +
                           "               c.U_cod , c.U_cod1, c.U_co2 , c.U_cod3, c.U_cod4 , c.U_cod5, " +
                           "               c.U_cod6, c.U_cod7, c.U_cod8, c.U_cod9, c.U_cod10, c.U_cod11 " +
                           "          from OINV i inner join OCRD c  on i.CardCode = c.CardCode " +
                           "         where i.DocTotal > i.PaidToDate " + // Con Deuda
                           "           and i.DocDueDate < GETDATE()  " + // Deuda Vencida
                           "           and c.GroupCode = 100 " +
                           "           and c.CardType  = 'C') spr " +
                           " where s.CardCode = spr.CardCode  " +
                           "    or s.CardCode = spr.U_cod   " +
                           "    or s.CardCode = spr.U_cod1  " +
                           "    or s.CardCode = spr.U_co2   " +
                           "    or s.CardCode = spr.U_cod3  " +
                           "    or s.CardCode = spr.U_cod4  " +
                           "    or s.CardCode = spr.U_cod5  " +
                           "    or s.CardCode = spr.U_cod6  " +
                           "    or s.CardCode = spr.U_cod7  " +
                           "    or s.CardCode = spr.U_cod8  " +
                           "    or s.CardCode = spr.U_cod9  " +
                           "    or s.CardCode = spr.U_cod10 " +
                           "    or s.CardCode = spr.U_cod11 ";

                    RutAux = "";
                    FRecordSet.DoQuery(oSql);
                    while (!FRecordSet.EoF)
                    {
                        Rut = (System.String)FRecordSet.Fields.Item("LicTradNum").Value;
                        Rut = Rut.Substring(0, Rut.Length - 2);
                        if (Rut == RutAux)
                        {
                            FRecordSet.MoveNext();
                            continue;
                        }
                        else
                            RutAux = Rut;
                        try
                        {
                            if (conexionMysql.AbrirConexion() == true)
                            {

                                // Sobre MySQL
                                oSql = "SELECT ACCESO_AUTORIZACION, ACCESO_CUPOS_VISITAS FROM acceso_llacolen " +
                                       " WHERE ACCESO_RUT = '" + Rut + "'";

                                QryCmd = new MySqlCommand(oSql, conexionMysql.conexionMysql);

                                MySqlDataReader Reader = QryCmd.ExecuteReader();

                                if (Reader.HasRows)
                                {
                                    while (Reader.Read())
                                    {
                                        if ((Reader.GetString(0).Trim() != "N") || (Reader.GetInt32(1) != 0))
                                        {
                                            oSql = "UPDATE acceso_llacolen " +
                                                   "   SET ACCESO_AUTORIZACION  = 'N'" +
                                                   "      ,ACCESO_CUPOS_VISITAS = 0 " +
                                                   " WHERE ACCESO_RUT = '" + Rut + "'";

                                            MySqlConnection conn = new MySqlConnection(getLink());
                                            conn.Open();
                                            MySqlCommand QryCmd2 = new MySqlCommand(oSql, conn);
                                            
                                            if (QryCmd2.ExecuteNonQuery() > 0)
                                                oLog.LogMsg("Socio bloqueado en MySQL, Rut: " + Rut, "A", "D");
                                            else
                                                oLog.LogMsg("Fallo bloqueo de cliente sobre MySQL, Rut: " + Rut + " ************ ERROR ", "A", "D");

                                            conn.Close();
                                        }
                                    }
                                    Reader.Close();
                                }
                                conexionMysql.CerrarConexion();

                                // Sobre SBO
                                if (((System.String)FRecordSet.Fields.Item("frozenFor").Value) != "Y" || ((System.Int32)FRecordSet.Fields.Item("U_Tickets").Value != 0))
                                {
                                    if (oBP.GetByKey((System.String)FRecordSet.Fields.Item("CardCode").Value))
                                    {
                                        oBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES;
                                        oBP.Valid = SAPbobsCOM.BoYesNoEnum.tNO;
                                        oBP.UserFields.Fields.Item("U_Tickets").Value = 0;
                                        nErr = oBP.Update();
                                        if (nErr != 0)
                                        {
                                            sErr = oCompany.GetLastErrorDescription();
                                            oLog.LogMsg("Error al bloquear por pago en SBO, socio: " + (System.String)FRecordSet.Fields.Item("CardCode").Value + " - " + sErr, "A", "D");
                                        }
                                        else
                                            oLog.LogMsg("Bloqueo por pago en SBO, socio: " + (System.String)FRecordSet.Fields.Item("CardCode").Value, "A", "D");
                                    }
                                }
                                FRecordSet.MoveNext();
                            }
                        }
                        catch (MySqlException ex)
                        {
                            oLog.LogMsg("MYSQLException " + ex.Message, "A", "D");
                            conexionMysql.CerrarConexion();
                            oLog.LogMsg("Cerrando Conexion EXCEPTION", "F", "E");
                        }
                    }

                    // Con Pago reciente => Desbloquear
                    oSql = "Select DISTINCT spr.cardcode SocPrin, s.CardCode, s.LicTradNum, s.GroupCode, s.frozenFor, s.U_Ausente, spr.saldo, spr.DocDueDate  " +
                           "       ,isnull((select GroupCode from OCRD where CardCode = s.U_cod1), -1) Grp102 " +
                           "  from OCRD s, " +
                           "       (select i.CardCode, i.DocTotal-i.PaidToDate as Saldo, i.DocDueDate, " +
                           "               c.U_cod , c.U_cod1, c.U_co2 , c.U_cod3, c.U_cod4 , c.U_cod5, " +
                           "               c.U_cod6, c.U_cod7, c.U_cod8, c.U_cod9, c.U_cod10, c.U_cod11 " +
                           "          from OINV i inner join OCRD c  on i.CardCode = c.CardCode " +
                           "         where i.DocTotal = i.PaidToDate " + // Sin Deuda
                           "           and i.DocDueDate <= GETDATE() " + // Deuda Vencida
                           "           and i.DocDueDate > IsNull(U_SincPago, '1900-01-01') " + // Vencimiento > fecha sincroniz. => no actulizado
                           "           and c.GroupCode = 100 " +
                           "           and c.CardType  = 'C') spr " +
                           " where s.CardCode = spr.CardCode  " +
                           "    or s.CardCode = spr.U_cod   " +
                           "    or s.CardCode = spr.U_cod1  " +
                           "    or s.CardCode = spr.U_co2   " +
                           "    or s.CardCode = spr.U_cod3  " +
                           "    or s.CardCode = spr.U_cod4  " +
                           "    or s.CardCode = spr.U_cod5  " +
                           "    or s.CardCode = spr.U_cod6  " +
                           "    or s.CardCode = spr.U_cod7  " +
                           "    or s.CardCode = spr.U_cod8  " +
                           "    or s.CardCode = spr.U_cod9  " +
                           "    or s.CardCode = spr.U_cod10 " +
                           "    or s.CardCode = spr.U_cod11 ";

                    RutAux = "";
                    FRecordSet.DoQuery(oSql);
                    while (!FRecordSet.EoF)
                    {
                        Autorizacion = "Y";
                        if ((System.String)FRecordSet.Fields.Item("U_Ausente").Value == "2")  // U_Ausente = 2 => Si, 1 => No
                            Autorizacion = "N";
                        if ((System.String)FRecordSet.Fields.Item("FrozenFor").Value == "Y")
                            Autorizacion = "N";

                        Rut = (System.String)FRecordSet.Fields.Item("LicTradNum").Value;
                        Rut = Rut.Substring(0, Rut.Length - 2);

                        Tickets = "0";
                        if (((System.Int32)FRecordSet.Fields.Item("GroupCode").Value == 100) && ((System.Int32)FRecordSet.Fields.Item("Grp102").Value == -1))
                            Tickets = "20";
                        else if (((System.Int32)FRecordSet.Fields.Item("GroupCode").Value == 100) || ((System.Int32)FRecordSet.Fields.Item("GroupCode").Value == 102))
                            Tickets = "10";
                        //MYSQL 
                        if (conexionMysql.AbrirConexion() == true)
                        {
                            oSql = "UPDATE acceso_llacolen " +
                                   "   SET ACCESO_CUPOS_VISITAS = " + Tickets +
                                   "      ,ACCESO_AUTORIZACION  = '" + Autorizacion + "'" +
                                   " WHERE ACCESO_RUT = '" + Rut + "'";
                            QryCmd = new MySqlCommand(oSql, conexionMysql.conexionMysql);
                            //QryCmd = new MySqlCommand(oSql, FBDBatchMySQL);

                            if (QryCmd.ExecuteNonQuery() > 0)
                                oLog.LogMsg("Socio actualizado, " + Tickets + " cupos asignados, en MySQL, RUT: " + Rut, "A", "D");
                            else
                                oLog.LogMsg("Fallo Actualizacion de socio en MySQL: " + Rut + " ************ ERROR ", "A", "D");

                            conexionMysql.CerrarConexion();

                            if (oBP.GetByKey((System.String)FRecordSet.Fields.Item("CardCode").Value))
                            {
                                oBP.Frozen = SAPbobsCOM.BoYesNoEnum.tNO;
                                oBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES;
                                oBP.UserFields.Fields.Item("U_Tickets").Value = Int32.Parse(Tickets);
                                oBP.UserFields.Fields.Item("U_SincPago").Value = (System.DateTime)FRecordSet.Fields.Item("DocDueDate").Value;
                                nErr = oBP.Update();
                                if (nErr != 0)
                                {
                                    sErr = oCompany.GetLastErrorDescription();
                                    oLog.LogMsg("Error al actualizar por pago en SBO, socio: " + (System.String)FRecordSet.Fields.Item("CardCode").Value + " - " + sErr, "A", "D");
                                }
                                else
                                    oLog.LogMsg("Actualización por pago en SBO, socio: " + (System.String)FRecordSet.Fields.Item("CardCode").Value, "A", "D");
                            }
                        }
                        FRecordSet.MoveNext();
                    }
                
            }
            catch (Exception e)
            {
                oLog.LogMsg("Error al sincronizar SBO : " + e.Message, "A", "E");
            }
        }
    }
}
