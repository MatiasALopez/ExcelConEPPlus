using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;

namespace Excel
{
    /// <summary>
    /// Clase base que representa un archivo Excel.
    /// </summary>
    public abstract class ExcelBase
    {
        public ExcelBase()
        {
            Errores = new List<string>();
        }

        #region Propiedades

        /// <summary>
        /// Lista de errores encontrados.
        /// </summary>
        public List<string> Errores { get; private set; }

        #endregion Propiedades
        
        #region Metodos

        #region Publicos

        /// <summary>
        /// Lee el contenido de cada una de las hojas seleccionadas.
        /// Al finalizar la lectura, cada hoja contendrá los registros y errores encontrados.
        /// En caso de producirse algún error relacionado con el proceso de lectura del archivo, se lo guardará en la propiedad Errores.
        /// </summary>
        /// <param name="bytes"></param>
        public void Leer(byte[] bytes)
        {
            try
            {
                using (var stream = new MemoryStream(bytes))
                {
                    LeerRegistrosDeHojas(stream);
                }
            }
            catch (Exception ex)
            {
                Errores.Add(ex.Message);
            }
        }

        /// <summary>
        /// Lee el contenido del Excel.
        /// </summary>
        /// <param name="path"></param>
        public void Leer(string path)
        {
            try
            {
                using (var stream = File.OpenRead(path))
                {
                    LeerRegistrosDeHojas(stream);
                }
            }
            catch (Exception ex)
            {
                Errores.Add(ex.Message);
            }
        }

        /// <summary>
        /// Lee el contenido del Excel.
        /// </summary>
        /// <param name="stream"></param>
        public void Leer(Stream stream)
        {
            try
            {
                LeerRegistrosDeHojas(stream);
            }
            catch (Exception ex)
            {
                Errores.Add(ex.Message);
            }
        }

        #endregion Publicos

        #region Protegidos

        /// <summary>
        /// Obtiene las hojas que se desean leer.
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        protected abstract IEnumerable<IExcelHoja> ObtenerHojas(ExcelWorkbook workbook);

        #endregion Protegidos

        #region Privados

        /// <summary>
        /// Lee los registros de las hojas.
        /// </summary>
        /// <param name="stream"></param>
        private void LeerRegistrosDeHojas(Stream stream)
        {
            using (var package = new ExcelPackage(stream))
            {
                var hojas = ObtenerHojas(package.Workbook);
                foreach (var hoja in hojas)
                {
                    try
                    {
                        hoja.LeerRegistros();
                    }
                    catch (Exception ex)
                    {
                        Errores.Add(ex.Message);
                    }
                }
            }
        }

        #endregion Privados

        #endregion Metodos
    }
}
