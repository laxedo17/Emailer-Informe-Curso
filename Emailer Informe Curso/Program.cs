using Emailer_Informe_Curso.Modelos;
using Emailer_Informe_Curso.Repository;
using Emailer_Informe_Curso.Trabajadores;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Data;

namespace Emailer_Informe_Curso
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ComandoInformeDetalladoMatriculas comando = new ComandoInformeDetalladoMatriculas(@"Data Source=localhost;Initial Catalog=InformeCurso;Integrated Security=True");//dentro vai o nome da conexion que obtemos pulsando en Propiedades da base de datos no Server Explorer
                IList<ModeloInformeDetalladoMatriculas> modelos = comando.ObtenerLista();

                var nombreFicheroInforme = "InformeDetallesMatriculas.xlsx";
                CreadorHojaExcelInformeDetalladoMatriculas creadorHojaMatriculas = new CreadorHojaExcelInformeDetalladoMatriculas();
                creadorHojaMatriculas.Crear(nombreFicheroInforme, modelos);

                EnviaEmailerInformeDetalleMatriculas emailer = new EnviaEmailerInformeDetalleMatriculas();
                emailer.Enviar(nombreFicheroInforme);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Algo fue mal: {0}", ex.Message);
            }



            //    ModeloInformeDetalladoMatriculas modelo = new ModeloInformeDetalladoMatriculas()
            //    {
            //        IdMatricula = 1,
            //        Nombre = "Lino",
            //        Apellido = "Lana",
            //        CodigoCurso = "AN",
            //        Descripcion = "descripcion"
            //    };

            //Parte de EXEMPLO
            /*var json = JsonConvert.SerializeObject(modelo); *///transforma o obxeto de arriba a formato de obxeto Json

            /*ModeloInformeDetalladoMatriculas objetoDeJson = (ModeloInformeDetalladoMatriculas)JsonConvert.DeserializeObject(json, typeof(ModeloInformeDetalladoMatriculas));*/ //transforma un obxeto Json de novo a formato obxeto
                                                                                                                                                                                 //-------FIN parte de EXEMPLO--------

            //DataTable tabla = new DataTable();

            //DataColumn columna1 = new DataColumn("IdMatricula", typeof(int));

            //DataColumn columna2 = new DataColumn("Nombre", typeof(string));

            //DataColumn columna3 = new DataColumn("Apellido", typeof(string));

            //DataColumn columna4 = new DataColumn("CodigoCurso", typeof(string));

            //DataColumn columna5 = new DataColumn("Descripcion", typeof(string));

            //tabla.Columns.Add(columna1);
            //tabla.Columns.Add(columna2);
            //tabla.Columns.Add(columna3);
            //tabla.Columns.Add(columna4);
            //tabla.Columns.Add(columna5);

            //tabla.Rows.Add(1, "Lino", "Lana", "An", "descripcion");

            //foreach (DataRow fila in tabla.Rows)
            //{
            //    foreach(DataColumn columna in tabla.Columns)
            //    {
            //        Console.WriteLine(fila[columna]); //imprimimos de cada fila as columnas -IdMatricula, Nombre, Apellido, CodigoCurso, Descripcion- 1 a 1 cos valores que temos na taboa
        }
    }
}

