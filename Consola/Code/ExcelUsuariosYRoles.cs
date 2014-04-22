using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;

using Excel;

namespace ExcelConEPPlus
{
    public class ExcelUsuariosYRoles : ExcelBase
    {
        #region Propiedades

        public HojaUsuarios HojaUsuarios { get; private set; }
        public HojaRoles HojaRoles { get; private set; }
        public HojaUsuariosRoles HojaUsuariosRoles { get; private set; }
        
        #endregion Propiedades
        
        #region Metodos

        #region Protegidos

        protected override IEnumerable<IExcelHoja> ObtenerHojas(ExcelWorkbook workbook)
        {
            HojaUsuarios = new HojaUsuarios(workbook);
            HojaRoles = new HojaRoles(workbook);
            HojaUsuariosRoles = new HojaUsuariosRoles(workbook);

            return new List<IExcelHoja> { HojaUsuarios, HojaRoles, HojaUsuariosRoles };
        }

        #endregion Protegidos

        #endregion Metodos
    }

    public class RegistroUsuario : ExcelRegistroBase
    {
        public RegistroUsuario(ExcelWorksheet hoja, int fila)
            : base(hoja, fila)
        {
            NombreDeUsuario = ObtenerTexto(1, "Nombre de usuario");
            NombreCompleto = ObtenerTexto(2, "Nombre completo");
            FechaNacimiento = ObtenerValor<DateTime>(3, "Fecha de nacimiento");
            Categoria = ObtenerEnum<UsuarioCategoria>(4, "Categoría");
            EstaActivo = ObtenerValor<bool>(5, "Está activo");
            FechaDeBloqueo = ObtenerValor<DateTime>(6, "Fecha de bloqueo", esRequerido: false);
            Comentarios = ObtenerTexto(7, "Comentarios", esRequerido: false);
        }

        public string NombreDeUsuario { get; set; }
        public string NombreCompleto { get; set; }
        public DateTime? FechaNacimiento { get; set; }
        public UsuarioCategoria? Categoria { get; set; }
        public bool? EstaActivo { get; set; }
        public DateTime? FechaDeBloqueo { get; set; }
        public string Comentarios { get; set; }

        public override bool EsRegistroVacio
        {
            get
            {
                return
                    string.IsNullOrWhiteSpace(NombreDeUsuario) &&
                    string.IsNullOrWhiteSpace(NombreCompleto) &&
                    !FechaNacimiento.HasValue &&
                    !Categoria.HasValue &&
                    !EstaActivo.HasValue &&
                    !FechaDeBloqueo.HasValue &&
                    string.IsNullOrWhiteSpace(Comentarios);
            }
        }
    }

    public class RegistroRol : ExcelRegistroBase
    {
        public RegistroRol(ExcelWorksheet hoja, int fila)
            : base(hoja, fila)
        {
            NombreDeRol = ObtenerTexto(1, "Nombre de rol");
        }

        public string NombreDeRol { get; set; }

        public override bool EsRegistroVacio
        {
            get
            {
                return
                    string.IsNullOrWhiteSpace(NombreDeRol);
            }
        }
    }

    public class RegistroUsuarioRol : ExcelRegistroBase
    {
        public RegistroUsuarioRol(ExcelWorksheet hoja, int fila)
            : base(hoja, fila)
        {
            NombreDeUsuario = ObtenerTexto(1, "Nombre de usuario");
            NombreDeRol = ObtenerTexto(2, "Nombre de rol");
        }

        public string NombreDeUsuario { get; set; }
        public string NombreDeRol { get; set; }

        public override bool EsRegistroVacio
        {
            get
            {
                return
                    string.IsNullOrWhiteSpace(NombreDeUsuario) &&
                    string.IsNullOrWhiteSpace(NombreDeRol);
            }
        }
    }

    public class HojaUsuarios : ExcelHojaBase<RegistroUsuario>
    {
        public HojaUsuarios(ExcelWorkbook workbook)
            : base(workbook,
                nombre: "Usuarios", 
                encabezados: new string[] { "Nombre de usuario", "Nombre completo", "Fecha de nacimiento", "Categoria", "Esta activo", "Fecha de bloqueo", "Comentarios" },
                filaEncabezados: 2, 
                filaPrimerRegistro: 3)
        { }

        protected override RegistroUsuario InstanciarRegistro(int fila)
        {
            return new RegistroUsuario(Worksheet, fila);
        }
    }

    public class HojaRoles : ExcelHojaBase<RegistroRol>
    {
        public HojaRoles(ExcelWorkbook workbook)
            : base(workbook, 
                nombre: "Roles",
                encabezados: new string[] { "Nombre de rol" })
        { }

        protected override RegistroRol InstanciarRegistro(int fila)
        {
            return new RegistroRol(Worksheet, fila);
        }
    }

    public class HojaUsuariosRoles : ExcelHojaBase<RegistroUsuarioRol>
    {
        public HojaUsuariosRoles(ExcelWorkbook workbook)
            : base(workbook, 
                nombre: "UsuariosRoles", 
                encabezados: new string[] { "Nombre de usuario", "Nombre de rol" })
        { }

        protected override RegistroUsuarioRol InstanciarRegistro(int fila)
        {
            return new RegistroUsuarioRol(Worksheet, fila);
        }
    }

    public enum UsuarioCategoria
    {
        Junior = 1,
        SemiSenior = 2,
        Senior = 3,
        Especialista = 4
    }
}
