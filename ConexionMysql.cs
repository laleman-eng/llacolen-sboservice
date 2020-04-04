using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace Llacolen_SBOService
{
    public class ConexionMysql
    {
        public MySqlConnection conexionMysql;

        private Logs.Logger FLog;
        public Logs.Logger oLog
        {
            get { return this.FLog; }
            set { this.FLog = value; }
        }


        public ConexionMysql()
        {
           string stringConnection =     "Server    =" + Llacolen_SBOService.Properties.Settings.Default.BatchServer +
                                         ";port     =" + Llacolen_SBOService.Properties.Settings.Default.BatchMySQLPort +
                                         ";Database =" + Llacolen_SBOService.Properties.Settings.Default.BatchDB +
                                         ";Uid      =" + Llacolen_SBOService.Properties.Settings.Default.BatchUser +
                                         ";Password =" + Llacolen_SBOService.Properties.Settings.Default.BatchPassword +
                                         "";
            conexionMysql = new MySqlConnection(stringConnection);
        }

        public bool AbrirConexion()
        {
            try
            {
                conexionMysql.Open();
               // oLog.LogMsg("Connected to MYSQL 2", "F", "D");
                return true;
            }
            catch (MySqlException ex )
            {
               oLog.LogMsg("Excepcion to Open MYSQL "+ex.Message, "F", "E");   
                return false;
            }
            
        }

        public bool CerrarConexion()
        {
            try
            {
                conexionMysql.Close();
               // oLog.LogMsg("Disconnected to MYSQL ", "F", "D");
                return true;
            }
            catch (MySqlException ex)
            {
                oLog.LogMsg("Excepcion to Disconnected MYSQL " + ex.Message, "F", "E"); 
                return false;
            }
        }
    }
}
