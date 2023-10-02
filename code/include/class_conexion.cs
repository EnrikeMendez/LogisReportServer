using System;
using System.Configuration;

namespace ReportServer2022
{
    public class class_conexion
    {
        //public String HOST;
        //public String PORT;
        //public String DB;
        //public String PASSWORD;
        //public String USER_ID;

        public string db_cadena;

        public void sub_set_conexion()
        {

            //HOST = "192.168.0.4";
            //PORT = "1521";
            //DB = "Orfeo2";
            //USER_ID = "web_adm";
            //PASSWORD = "va4ncMC3P";

            //DESARROLLO:
            //HOST = "192.168.0.130";
            //PORT = "1521";
            //DB = "DEVORFEO";
            //USER_ID = "web_adm";
            //PASSWORD = "va4ncMC3P";

            //db_cadena = "DATA SOURCE = "+ HOST +":"+PORT+" / "+DB+"; PASSWORD = "+ PASSWORD +"; USER ID = "+ USER_ID +";";
            db_cadena = ConfigurationSettings.AppSettings["dbConnection"].ToString();
        }
        /// <summary>
        /// Libera los recursos utilizados por el objeto.
        /// </summary>
        public void Dispose()
        {
            if (db_cadena != null)
            {
                db_cadena = string.Empty;
                GC.SuppressFinalize(db_cadena);
            }
        }
    }
}