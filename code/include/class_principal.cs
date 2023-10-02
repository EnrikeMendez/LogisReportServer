using ReportServer2022.code.include;
using ReportServer2022.code.querys;
using System;

namespace ReportServer2022
{
    public class class_principal
    {
        public class_querys obj_querys = new class_querys();
        public class_xfunciones obj_xfunciones = new class_xfunciones();

        /// <summary>
        /// Libera los recursos utilizados por el objeto.
        /// </summary>
        public void Dispose()
        {
            if (this.obj_querys != null)
            {
                GC.SuppressFinalize(this.obj_querys);
            }
            if (this.obj_xfunciones != null)
            {
                GC.SuppressFinalize(this.obj_xfunciones);
            }
            GC.Collect();
        }
    }
}