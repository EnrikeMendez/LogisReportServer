using System;
using System.Data;
using System.Data.OracleClient;

namespace ReportServer2022
{
    public class class_llena_tabla
    {
        public string cadena_conexion;
        public string sqltext;
        public string error_datos;
        public DataTable tabla_llena;

        public void sub_llenar_tabla()
        {
            OracleConnection conexion = null;
            OracleDataAdapter Adapter3 = null;

            try
            {
                conexion = new OracleConnection();
                tabla_llena = new DataTable();
                tabla_llena.Clear();
                conexion.ConnectionString = cadena_conexion;

                conexion.Open();

                Adapter3 = new OracleDataAdapter(sqltext, conexion);
                Adapter3.Fill(tabla_llena);

                conexion.Close();
            }
            catch (Exception ex)
            {
                error_datos = ex + "" + " \n " + sqltext;
            }
            finally
            {
                if (Adapter3 != null)
                {
                    Adapter3.Dispose();
                    GC.SuppressFinalize(Adapter3);
                }
                if (conexion != null)
                {
                    if (conexion.State == ConnectionState.Open)
                    {
                        conexion.Close();
                    }
                    GC.SuppressFinalize(conexion);
                }
                GC.Collect();
            }
        }
        /// <summary>
        /// Libera los recursos utilizados por el objeto.
        /// </summary>
        public void Dispose()
        {
            if (cadena_conexion != null)
            {
                cadena_conexion = string.Empty;
                cadena_conexion = null;
            }
            if (sqltext != null)
            {
                sqltext = string.Empty;
                sqltext = null;
            }
            if (error_datos != null)
            {
                error_datos = string.Empty;
                error_datos = null;
            }
            if (tabla_llena != null)
            {
                tabla_llena.Dispose();
                GC.SuppressFinalize(tabla_llena);
            }
            GC.Collect();
        }
    }
}