using Dapper;

using Emailer_Informe_Curso.Modelos;

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace Emailer_Informe_Curso.Repository
{
    class ComandoInformeDetalladoMatriculas
    {
        private string _stringConexion;

        public ComandoInformeDetalladoMatriculas(string stringConexion)
        {
            _stringConexion = stringConexion;
        }

        public IList<ModeloInformeDetalladoMatriculas> ObtenerLista()
        {
            List<ModeloInformeDetalladoMatriculas> informeDetalladoMatricula = new List<ModeloInformeDetalladoMatriculas>();

            var sql = "InformeCurso_ObtenerLista";

            using (SqlConnection conexion = new SqlConnection(_stringConexion))
            {
                informeDetalladoMatricula = conexion.Query<ModeloInformeDetalladoMatriculas>(sql).ToList();
            }

            return informeDetalladoMatricula;
        }
    }
}
