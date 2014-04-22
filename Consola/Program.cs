using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConEPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var excel = new ExcelUsuariosYRoles();
                
                //  Leer excel  (usando los bytes del archivo)
                excel.Leer(System.IO.File.ReadAllBytes(@"Archivos\Usuarios y Roles.xlsx"));

                //  Leer excel CON errores (usando el path del archivo)
                excel.Leer(@"Archivos\Usuarios y Roles (con errores).xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("[ERROR] {0}", ex.Message));
            }
        }
    }
}
