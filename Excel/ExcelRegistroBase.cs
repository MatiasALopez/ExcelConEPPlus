using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Excel
{
    /// <summary>
    /// Clase base que representa un registro de una hoja de un archivo Excel.
    /// </summary>
    public abstract class ExcelRegistroBase
    {
        public ExcelRegistroBase(ExcelWorksheet hoja, int fila)
        {
            if (hoja == null)
                throw new ArgumentNullException("hoja");

            Hoja = hoja;
            Fila = fila;
            Errores = new List<string>();
        }

        #region Propiedades

        #region Publicas

        public int Fila { get; private set; }
        public List<string> Errores { get; private set; }
        public bool EsValido
        {
            get { return Errores.Count == 0; }
        }

        /// <summary>
        /// Indica si el registro es un registro vacío.
        /// </summary>
        public abstract bool EsRegistroVacio { get; }

        #endregion Publicos

        #region Privadas

        private ExcelWorksheet Hoja { get; set; }

        #endregion Privadas

        #endregion Propiedades

        #region Metodos

        #region Protegidos

        protected string ObtenerTexto(int columna, string descripcion, bool esRequerido = true)
        {
            var celda = Hoja.Cells[Fila, columna];
            var valor = celda.Value;

            if (valor != null)
                return Convert.ToString(valor);
            else
            {
                if (esRequerido)
                    Errores.Add(string.Format("El dato '{0}' es requerido y no ha sido especificado (celda {1}).", descripcion, celda.Address));

                return null;
            }
        }

        protected T? ObtenerValor<T>(int columna, string descripcion, bool esRequerido = true)
            where T : struct
        {
            var celda = Hoja.Cells[Fila, columna];
            var valor = celda.Value;

            if (valor != null)
            {
                try
                {
                    return (T)Convert.ChangeType(valor, typeof(T));
                }
                catch (Exception)
                {
                    Errores.Add(string.Format("El dato '{0}' no es válido (celda {1}).", descripcion, celda.Address));
                    return null;
                }
            }
            else
            {
                if (esRequerido)
                    Errores.Add(string.Format("El dato '{0}' es requerido y no ha sido especificado (celda {1}).", descripcion, celda.Address));

                return null;
            }
        }

        public T? ObtenerEnum<T>(int columna, string descripcion, bool esRequerido = true)
            where T : struct
        {
            var tipo = typeof(T);
            if (!tipo.IsEnum)
                throw new ArgumentException(string.Format("El tipo '{0}' no es un Enum válido.", tipo.Name));

            var celda = Hoja.Cells[Fila, columna];
            var valor = celda.Value;

            if (valor != null)
            {
                try
                {
                    return (T?)Enum.Parse(tipo, Convert.ToString(valor));
                }
                catch (Exception)
                {
                    Errores.Add(string.Format("El dato '{0}' no es válido (celda {1}).", descripcion, celda.Address));
                    return null;
                }
            }
            else
            {
                if (esRequerido)
                    Errores.Add(string.Format("El dato '{0}' es requerido y no ha sido especificado (celda {1}).", descripcion, celda.Address));

                return null;
            }
        }

        #endregion Protegidos

        #endregion Metodos
    }
}
