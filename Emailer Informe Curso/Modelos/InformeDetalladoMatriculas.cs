using System;
using System.Collections.Generic;
using System.Text;

namespace Emailer_Informe_Curso.Modelos
{
    class InformeDetalladoMatriculas
    {
        public int IdMatricula { get; set; }
        public string Nombre { get; set; }
        public string Apellido { get; set; }
        public string CodigoCurso { get; set; }
        public string Descripcion { get; set; }
    }
}
