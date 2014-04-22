using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;

namespace Excel
{
    public interface IExcelHoja
    {
        void LeerRegistros();
    }

    /// <summary>
    /// Clase base que representa una hoja de un archivo Excel.
    /// </summary>
    /// <typeparam name="TRegistro">Tipo del registro que contiene la hoja.</typeparam>
    public abstract class ExcelHojaBase<TRegistro> : IExcelHoja
        where TRegistro : ExcelRegistroBase
    {
        /// <summary>
        /// Crea una nueva hoja.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="nombre">Nombre de la hoja.</param>
        /// <param name="encabezados">Lista de encabezados.</param>
        /// <param name="filaEncabezados">Fila en la cual se encuentran los encabezados. Por default, es 1.</param>
        /// <param name="filaPrimerRegistro">Fila en la cual se encuentra el primer registro. Por default, es 2.</param>
        public ExcelHojaBase(ExcelWorkbook workbook, string nombre, IEnumerable<string> encabezados, int filaEncabezados = 1, int filaPrimerRegistro = 2)
        {
            if (workbook == null)
                throw new ArgumentNullException("workbook");
            if (string.IsNullOrWhiteSpace(nombre))
                throw new ArgumentNullException("nombre");
            if (encabezados == null || encabezados.Count() == 0)
                throw new ArgumentNullException("encabezados");

            Workbook = workbook;
            Nombre = nombre;

            Encabezados = encabezados;
            FilaEncabezados = filaEncabezados;
            FilaPrimerRegistro = filaPrimerRegistro;

            Registros = new List<TRegistro>();
            Errores = new List<string>();

            Worksheet = Workbook.Worksheets[Nombre];
            if (Worksheet == null)
                Errores.Add(string.Format("La hoja '{0}' no existe.", Nombre));
        }

        #region Propiedades

        #region Publicas

        public string Nombre { get; private set; }

        public IEnumerable<string> Encabezados { get; private set; }
        public int FilaEncabezados { get; private set; }
        public int FilaPrimerRegistro { get; private set; }

        public List<TRegistro> Registros { get; private set; }
        public List<string> Errores { get; private set; }

        #endregion Publicas

        #region Protegidas

        protected ExcelWorksheet Worksheet { get; private set; }

        #endregion Protegidas

        #region Privadas

        private ExcelWorkbook Workbook { get; set; }

        #endregion Privadas
        
        #endregion Propiedades
        
        #region Metodos

        #region Publicos

        /// <summary>
        /// Lee los registros de la hoja. 
        /// Asume que el último registro es el inmediato anterior a una fila vacía.
        /// Al finalizar la lectura, se habrán cargado los registros y errores encontrados.
        /// </summary>
        public void LeerRegistros()
        {
            try
            {
                if (Worksheet == null)
                    return;

                Registros = new List<TRegistro>();
                Errores = new List<string>();

                if (!ValidarEncabezados())
                    return;

                for (int fila = FilaPrimerRegistro; ; fila++)
                {
                    var registro = InstanciarRegistro(fila);
                    if (registro.EsRegistroVacio)
                        break;

                    if (!registro.EsValido)
                    {
                        Errores.Add(string.Format("El registro de la fila {1} tiene errores.{0}{2}", Environment.NewLine, fila, string.Join(Environment.NewLine, registro.Errores)));
                        continue;
                    }

                    Registros.Add(registro);
                }
            }
            catch (Exception ex)
            {
                Errores.Add(ex.Message);
            }
        }

        #endregion Publicos

        #region Privados

        /// <summary>
        /// Valida los encabezados de la hoja.
        /// </summary>
        /// <returns></returns>
        private bool ValidarEncabezados()
        {
            string encabezado;
            int columna;
            ExcelRange celda;
            int cantColumnas = Encabezados.Count();
            for (int i = 0; i < cantColumnas; i++)
            {
                encabezado = Encabezados.ElementAt(i);
                columna = i + 1;
                celda = Worksheet.Cells[FilaEncabezados, columna];

                if (!Convert.ToString(celda.Value).Equals(encabezado))
                    Errores.Add(string.Format("Encabezado '{0}' no encontrado (celda {1}).", encabezado, celda.Address));
            }

            return !Errores.Any();
        }

        #endregion Privados

        #region Protegidos

        /// <summary>
        /// Crea una instancia del registro de la fila especificada.
        /// </summary>
        /// <param name="fila"></param>
        /// <returns></returns>
        protected abstract TRegistro InstanciarRegistro(int fila);

        #endregion Protegidos

        #endregion Metodos
    }
}
