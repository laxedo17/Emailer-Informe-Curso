using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Emailer_Informe_Curso.Modelos;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Emailer_Informe_Curso.Trabajadores
{
    class CreadorHojaExcelInformeDetalladoMatriculas
    {
        public void Crear(string nombreArchivo, IList<ModeloInformeDetalladoMatriculas> modelosMatricula)
        {
            using (SpreadsheetDocument documento = SpreadsheetDocument.Create(nombreArchivo, SpreadsheetDocumentType.Workbook))
            {
                var json = JsonConvert.SerializeObject(modelosMatricula);

                //TRUCO para Json
                /*En vez de facer algo como esto
                 ModeloInformeDetalladoMatriculas objetoDeJson = (ModeloInformeDetalladoMatriculas)JsonConvert.DeserializeObject(json, typeof(ModeloInformeDetalladoMatriculas));

                Agora usamos DataTable directamente e non o tipo de obxeto, co cal transformamos directamente a unha DataTable
                 */
                DataTable tablaMatriculas = (DataTable)JsonConvert.DeserializeObject(json, typeof(DataTable)); //typeof e un datatype neste caso

                //Hay Sheet, que e como un componente pai en Excel. Sheet en Excel significa que a sheet pode ser calquer tipo, pode ser Chart Sheet, Work Sheet, etc. 
                //Un Workbook e o documento enteiro, e un Workbook pode ter diferentes Sheets (work sheet, chart sheet, etc)
                //Un WorkSheet

                //Primeiro temos que crear un Workbook
                WorkbookPart workbookPart = documento.AddWorkbookPart();//workbookpart e un obxeto que conten configuracions globales dos componentes dun Workdbook.
                workbookPart.Workbook = new Workbook();//Workbook crea o verdadeiro obxeto Workbook -diferente a WorkbookPart-

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(); //de ese workbookPart queremos engadir unha nova parte ao Workbook, e esa e a worksheetPart

                //SheetData son os datos reales que se van a ver na folla de Excel
                SheetData datosDeHojaExcel = new SheetData(); //neste obxeto e onde imos albergar as filas e columnas cos datos
                worksheetPart.Worksheet = new Worksheet(datosDeHojaExcel); //WorkSheet obten un novo obxeto chamado datosDeHojaExcel

                //necesitamos asociar a nosa Worksheetpart co noso Workbook, porque agora mismo inda son unha worksheetpart e workbook separados
                Sheets listaHojas = documento.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());//lista vacia de follas

                Sheet hojaSimple = new Sheet()
                {
                    Id = documento.WorkbookPart.GetIdOfPart(worksheetPart), //con esta Id automatica finalmente asociamos a Worksheetpart co noso Workbook, o que nos di esta instruccion e que este obxeto Sheet vai a ser a Worksheet (o obxeto worksheetPart neste caso) que usaremos para asociar o workbook
                    SheetId = 1,
                    Name = "Hoja Informe" //O nome da folla de Excel vai ser o indicado ahi
                };

                listaHojas.Append(hojaSimple);

                //imos a situar a nosa Datatable na folla de calculo
                Row filaEncabezadoExcel = new Row();

                foreach (DataColumn columnaTabla in tablaMatriculas.Columns)
                {
                    Cell celda = new Cell();
                    celda.DataType = CellValues.String;
                    celda.CellValue = new CellValue(columnaTabla.ColumnName); //isto vai a crear unha fila cos datos de cada columna
                    filaEncabezadoExcel.Append(celda);
                }

                datosDeHojaExcel.AppendChild(filaEncabezadoExcel);

                foreach (DataRow filaTabla in tablaMatriculas.Rows)
                {
                    //columna1 columna2 columna3
                    //datos1   datos1   datos1
                    //datos2   datos2   datos2
                    //asi quedarian os datos tras facer o loop de abaixo

                    Row nuevaFilaExcel = new Row(); //recolle os datos de cada columna de arriba
                    foreach (DataColumn columnaTabla in tablaMatriculas.Columns)
                    {
                        Cell celda = new Cell();
                        celda.DataType = CellValues.String;
                        celda.CellValue = new CellValue(filaTabla[columnaTabla.ColumnName].ToString());
                        nuevaFilaExcel.AppendChild(celda);
                    }
                    datosDeHojaExcel.AppendChild(nuevaFilaExcel);//cada vez que terminamos todo o de arriba, cada fila hai que situala en datos da folla de Excel, o cal facemos con esta instruccion
                }

                workbookPart.Workbook.Save();//garda o noso workbookpart

                //aqui poderiamos situar unha folla de graficos ou similar,et
                //Sheet hojaGraficos = new Sheet()
                //{
                //    Id = documento.WorkbookPart.GetIdOfPart(worksheetPart), //con esta Id automatica finalmente asociamos a Worksheetpart co noso Workbook, o que nos di esta instruccion e que este obxeto Sheet vai a ser a Worksheet (o obxeto worksheetPart neste caso) que usaremos para asociar o workbook
                //    SheetId = 1,
                //    Name = "Hoja Informe" //O nome da folla de Excel vai ser o indicado ahi
                //};
            }
        }
    }
}
