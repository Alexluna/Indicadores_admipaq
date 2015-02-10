using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Globalization;

namespace IndicadoresISEL.Modelo
{
    class Modelo_Impresion
    {






        #region IMPRT EXCEL CRU
        public void excel_importCRUs(Tipos_Datos_CRU.ListDatosCRU ListDocmuentos)
        { //importar datos en excel
            try
            {


                // creating Excel Application
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;


                
                // see the excel sheet behind the program
                app.Visible = true;
                // get the reference of first sheet. By default its name is Sheet1.
                // store its reference to worksheet
                worksheet = workbook.Sheets["Hoja1"];
                worksheet = workbook.ActiveSheet;
                #region formato
                #region facturs
                Microsoft.Office.Interop.Excel.Range formatRange;
                formatRange = worksheet.get_Range("A4", "s1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion
                #region facturas público
                formatRange = worksheet.get_Range("W4", "AO1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                #endregion

                #region facturas ol
                formatRange = worksheet.get_Range("AR4", "BJ1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                #endregion
                #region ACUMULADOS ABONOS
                formatRange = worksheet.get_Range("BO4", "CG1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion
                #region ABONOS público
                formatRange = worksheet.get_Range("CJ4", "DB1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                #endregion
                #region ABONOS 3 AGENTE
                formatRange = worksheet.get_Range("DD4", "DV1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                #endregion

                #region zONA CENTRO
                formatRange = worksheet.get_Range("DY4", "EP1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion

                #region zONA SUR
                formatRange = worksheet.get_Range("ES4", "FJ1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                #endregion

                #region zONA NORTE
                formatRange = worksheet.get_Range("FM4", "GD1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                #endregion


                #region COMPRAS ACUMULADAS
                formatRange = worksheet.get_Range("GH4", "GZ1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion


                #region COMPRAS aNJI
                formatRange = worksheet.get_Range("HC4", "HV1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                #endregion

                #region PAGOS ACUMULADOS
                formatRange = worksheet.get_Range("HZ4", "IR1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                #endregion

                #region PAGOS aNJI
                formatRange = worksheet.get_Range("IU4", "JN1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion

                #region PRESTAMOS
                formatRange = worksheet.get_Range("JQ4", "KI1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                #endregion


                #region INGRESO TRASPASO
                formatRange = worksheet.get_Range("KL4", "LD1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                #endregion


                #region INGRESO DEV. GARANTIA
                formatRange = worksheet.get_Range("LH4", "LS1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion

                #region INGRESO DEV. GARANTIA
                formatRange = worksheet.get_Range("LU4", "MF1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                #endregion



                #endregion
                // changing the name of active sheet



                worksheet.Name = "Admipaq";
                int Row = 4;
                //titulo
                #region facturas
                Row = 4; //inicia a escribir en la fila 4
                #region encabezados
                worksheet.Cells[2, 5] = "facturas acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 1] = "Fecha";
                worksheet.Cells[Row, 2] = "Serie";
                worksheet.Cells[Row, 3] = "Folio";
                worksheet.Cells[Row, 4] = "Nombre del agente";
                worksheet.Cells[Row, 5] = "Razon social";
                worksheet.Cells[Row, 6] = "Fecha de vencimiento";
                worksheet.Cells[Row, 7] = "RFC";
                worksheet.Cells[Row, 8] = "Subtotal";
                worksheet.Cells[Row, 9] = "IVA";
                worksheet.Cells[Row, 10] = "TOTAL";
                worksheet.Cells[Row, 11] = "Pendiente";
                worksheet.Cells[Row, 12] = "Texto Extra 3";
                worksheet.Cells[Row, 13] = "Afectado";
                worksheet.Cells[Row, 14] = "Impreso";
                worksheet.Cells[Row, 15] = "Cancelado";
                worksheet.Cells[Row, 16] = "Total de unidades";
                worksheet.Cells[Row, 17] = "Clasificacion cliente2";
                worksheet.Cells[Row, 18] = "Texto extra1";
                worksheet.Cells[Row, 19] = "Nombre del concepto";
                //formato del archivo de excel
                
                /*
                Microsoft.Office.Interop.Excel.Range formatRange;
                formatRange = worksheet.get_Range("A4", "s1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna
                
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                */
                //titulo 
                worksheet.Cells[2, 27] = "facturas publico";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 23] = "Fecha";
                worksheet.Cells[Row, 24] = "Serie";
                worksheet.Cells[Row, 25] = "Folio";
                worksheet.Cells[Row, 26] = "Nombre del agente";
                worksheet.Cells[Row, 27] = "Razon social";
                worksheet.Cells[Row, 28] = "Fecha de vencimiento";
                worksheet.Cells[Row, 29] = "RFC";
                worksheet.Cells[Row, 30] = "Subtotal";
                worksheet.Cells[Row, 31] = "IVA";
                worksheet.Cells[Row, 32] = "TOTAL";
                worksheet.Cells[Row, 33] = "Pendiente";
                worksheet.Cells[Row, 34] = "Texto Extra 3";
                worksheet.Cells[Row, 35] = "Afectado";
                worksheet.Cells[Row, 36] = "Impreso";
                worksheet.Cells[Row, 37] = "Cancelado";
                worksheet.Cells[Row, 38] = "Total de unidades";
                worksheet.Cells[Row, 39] = "Clasificacion cliente2";
                worksheet.Cells[Row, 40] = "Texto extra1";
                worksheet.Cells[Row, 41] = "Nombre del concepto";
                //titulo 
                worksheet.Cells[2, 48] = "Facturas ol";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 44] = "Fecha";
                worksheet.Cells[Row, 45] = "Serie";
                worksheet.Cells[Row, 46] = "Folio";
                worksheet.Cells[Row, 47] = "Nombre del agente";
                worksheet.Cells[Row, 48] = "Razon social";
                worksheet.Cells[Row, 49] = "Fecha de vencimiento";
                worksheet.Cells[Row, 50] = "RFC";
                worksheet.Cells[Row, 51] = "Subtotal";
                worksheet.Cells[Row, 52] = "IVA";
                worksheet.Cells[Row, 53] = "TOTAL";
                worksheet.Cells[Row, 54] = "Pendiente";
                worksheet.Cells[Row, 55] = "Texto Extra 3";
                worksheet.Cells[Row, 56] = "Afectado";
                worksheet.Cells[Row, 57] = "Impreso";
                worksheet.Cells[Row, 58] = "Cancelado";
                worksheet.Cells[Row, 59] = "Total de unidades";
                worksheet.Cells[Row, 60] = "Clasificacion cliente2";
                worksheet.Cells[Row, 61] = "Texto extra1";
                worksheet.Cells[Row, 62] = "Nombre del concepto";
                Row++;
                #endregion

                #region contenido
                float total = 0;
                for (int i = 0; i < ListDocmuentos.facturas.Count; i++)
                {
                    worksheet.Cells[Row, 1] = ListDocmuentos.facturas[i].Fecha;
                    worksheet.Cells[Row, 2] = ListDocmuentos.facturas[i].Serie;
                    worksheet.Cells[Row, 3] = ListDocmuentos.facturas[i].Folio;
                    worksheet.Cells[Row, 4] = ListDocmuentos.facturas[i].NombreAgente;
                    worksheet.Cells[Row, 5] = ListDocmuentos.facturas[i].RazonSocial;
                    worksheet.Cells[Row, 6] = ListDocmuentos.facturas[i].FechaVencimiento;
                    worksheet.Cells[Row, 7] = ListDocmuentos.facturas[i].RFC;
                    worksheet.Cells[Row, 8] = ListDocmuentos.facturas[i].Subtotal;
                    worksheet.Cells[Row, 9] = ListDocmuentos.facturas[i].IVA;
                    worksheet.Cells[Row, 10] = ListDocmuentos.facturas[i].Total;
                    worksheet.Cells[Row, 11] = ListDocmuentos.facturas[i].Pendiente;
                    worksheet.Cells[Row, 12] = ListDocmuentos.facturas[i].TextoExtra3;
                    worksheet.Cells[Row, 13] = ListDocmuentos.facturas[i].Afectado;
                    worksheet.Cells[Row, 14] = ListDocmuentos.facturas[i].Impreso;
                    worksheet.Cells[Row, 15] = ListDocmuentos.facturas[i].Cancelado;
                    worksheet.Cells[Row, 16] = ListDocmuentos.facturas[i].TotalUnidades;
                    worksheet.Cells[Row, 17] = ListDocmuentos.facturas[i].Clasificacion2;
                    worksheet.Cells[Row, 18] = ListDocmuentos.facturas[i].TextoExtra1;
                    worksheet.Cells[Row, 19] = ListDocmuentos.facturas[i].NombreConcepto;
                    if (ListDocmuentos.facturas[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas[i].Total;
                        
                    }
                    Row++;
                }
                worksheet.Cells[2, 10] = "$ " + total;

                total = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.facturas_rfc_publico.Count; i++)
                {
                    worksheet.Cells[Row, 23] = ListDocmuentos.facturas_rfc_publico[i].Fecha;
                    worksheet.Cells[Row, 24] = ListDocmuentos.facturas_rfc_publico[i].Serie;
                    worksheet.Cells[Row, 25] = ListDocmuentos.facturas_rfc_publico[i].Folio;
                    worksheet.Cells[Row, 26] = ListDocmuentos.facturas_rfc_publico[i].NombreAgente;
                    worksheet.Cells[Row, 27] = ListDocmuentos.facturas_rfc_publico[i].RazonSocial;
                    worksheet.Cells[Row, 28] = ListDocmuentos.facturas_rfc_publico[i].FechaVencimiento;
                    worksheet.Cells[Row, 29] = ListDocmuentos.facturas_rfc_publico[i].RFC;
                    worksheet.Cells[Row, 30] = ListDocmuentos.facturas_rfc_publico[i].Subtotal;
                    worksheet.Cells[Row, 31] = ListDocmuentos.facturas_rfc_publico[i].IVA;
                    worksheet.Cells[Row, 32] = ListDocmuentos.facturas_rfc_publico[i].Total;
                    worksheet.Cells[Row, 33] = ListDocmuentos.facturas_rfc_publico[i].Pendiente;
                    worksheet.Cells[Row, 34] = ListDocmuentos.facturas_rfc_publico[i].TextoExtra3;
                    worksheet.Cells[Row, 35] = ListDocmuentos.facturas_rfc_publico[i].Afectado;
                    worksheet.Cells[Row, 36] = ListDocmuentos.facturas_rfc_publico[i].Impreso;
                    worksheet.Cells[Row, 37] = ListDocmuentos.facturas_rfc_publico[i].Cancelado;
                    worksheet.Cells[Row, 38] = ListDocmuentos.facturas_rfc_publico[i].TotalUnidades;
                    worksheet.Cells[Row, 39] = ListDocmuentos.facturas_rfc_publico[i].Clasificacion2;
                    worksheet.Cells[Row, 40] = ListDocmuentos.facturas_rfc_publico[i].TextoExtra1;
                    worksheet.Cells[Row, 41] = ListDocmuentos.facturas_rfc_publico[i].NombreConcepto;
                    if (ListDocmuentos.facturas_rfc_publico[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas_rfc_publico[i].Total;
                        
                    }
                    Row++;
                }
                worksheet.Cells[2, 32] = "$ " + total;
                total = 0;
                Row = 5;
                for (int i = 0; i < ListDocmuentos.facturas_rfc_ol.Count; i++)
                {
                    worksheet.Cells[Row, 44] = ListDocmuentos.facturas_rfc_ol[i].Fecha;
                    worksheet.Cells[Row, 45] = ListDocmuentos.facturas_rfc_ol[i].Serie;
                    worksheet.Cells[Row, 46] = ListDocmuentos.facturas_rfc_ol[i].Folio;
                    worksheet.Cells[Row, 47] = ListDocmuentos.facturas_rfc_ol[i].NombreAgente;
                    worksheet.Cells[Row, 48] = ListDocmuentos.facturas_rfc_ol[i].RazonSocial;
                    worksheet.Cells[Row, 49] = ListDocmuentos.facturas_rfc_ol[i].FechaVencimiento;
                    worksheet.Cells[Row, 50] = ListDocmuentos.facturas_rfc_ol[i].RFC;
                    worksheet.Cells[Row, 51] = ListDocmuentos.facturas_rfc_ol[i].Subtotal;
                    worksheet.Cells[Row, 52] = ListDocmuentos.facturas_rfc_ol[i].IVA;
                    worksheet.Cells[Row, 53] = ListDocmuentos.facturas_rfc_ol[i].Total;
                    worksheet.Cells[Row, 54] = ListDocmuentos.facturas_rfc_ol[i].Pendiente;
                    worksheet.Cells[Row, 55] = ListDocmuentos.facturas_rfc_ol[i].TextoExtra3;
                    worksheet.Cells[Row, 56] = ListDocmuentos.facturas_rfc_ol[i].Afectado;
                    worksheet.Cells[Row, 57] = ListDocmuentos.facturas_rfc_ol[i].Impreso;
                    worksheet.Cells[Row, 58] = ListDocmuentos.facturas_rfc_ol[i].Cancelado;
                    worksheet.Cells[Row, 59] = ListDocmuentos.facturas_rfc_ol[i].TotalUnidades;
                    worksheet.Cells[Row, 60] = ListDocmuentos.facturas_rfc_ol[i].Clasificacion2;
                    worksheet.Cells[Row, 61] = ListDocmuentos.facturas_rfc_ol[i].TextoExtra1;
                    worksheet.Cells[Row, 62] = ListDocmuentos.facturas_rfc_ol[i].NombreConcepto;
                    if (ListDocmuentos.facturas_rfc_ol[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas_rfc_ol[i].Total;
                        
                    }
                    Row++;
                }
                worksheet.Cells[2, 53] = "$ " + total;
                #endregion



                #endregion

                #region Abonos
                Row = 4;
                #region encabezados
                // worksheet.Cells[1, 1] = "Desglose de facturas";
                worksheet.Cells[2, 71] = "Acumulado de abonos";

                //encabezados facturas
                worksheet.Cells[Row, 67] = "Fecha";
                worksheet.Cells[Row, 68] = "Serie";
                worksheet.Cells[Row, 69] = "Folio";
                worksheet.Cells[Row, 70] = "Nombre del agente";
                worksheet.Cells[Row, 71] = "Razon social";
                worksheet.Cells[Row, 72] = "Fecha de vencimiento";
                worksheet.Cells[Row, 73] = "RFC";
                worksheet.Cells[Row, 74] = "Subtotal";
                worksheet.Cells[Row, 75] = "IVA";
                worksheet.Cells[Row, 76] = "TOTAL";
                worksheet.Cells[Row, 77] = "Pendiente";
                worksheet.Cells[Row, 78] = "Texto Extra 3";
                worksheet.Cells[Row, 79] = "Afectado";
                worksheet.Cells[Row, 80] = "Impreso";
                worksheet.Cells[Row, 81] = "Cancelado";
                worksheet.Cells[Row, 82] = "Total de unidades";
                worksheet.Cells[Row, 83] = "Clasificacion cliente2";
                worksheet.Cells[Row, 84] = "Texto extra1";
                worksheet.Cells[Row, 85] = "Nombre del concepto";
                //titulo 
                worksheet.Cells[2, 92] = "Abonos publico";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 88] = "Fecha";
                worksheet.Cells[Row, 89] = "Serie";
                worksheet.Cells[Row, 90] = "Folio";
                worksheet.Cells[Row, 91] = "Nombre del agente";
                worksheet.Cells[Row, 92] = "Razon social";
                worksheet.Cells[Row, 93] = "Fecha de vencimiento";
                worksheet.Cells[Row, 94] = "RFC";
                worksheet.Cells[Row, 95] = "Subtotal";
                worksheet.Cells[Row, 96] = "IVA";
                worksheet.Cells[Row, 97] = "TOTAL";
                worksheet.Cells[Row, 98] = "Pendiente";
                worksheet.Cells[Row, 99] = "Texto Extra 3";
                worksheet.Cells[Row, 100] = "Afectado";
                worksheet.Cells[Row, 101] = "Impreso";
                worksheet.Cells[Row, 102] = "Cancelado";
                worksheet.Cells[Row, 103] = "Total de unidades";
                worksheet.Cells[Row, 104] = "Clasificacion cliente2";
                worksheet.Cells[Row, 105] = "Texto extra1";
                worksheet.Cells[Row, 106] = "Nombre del concepto";

                //titulo 
                worksheet.Cells[2, 112] = "Abonos OL";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 108] = "Fecha";
                worksheet.Cells[Row, 109] = "Serie";
                worksheet.Cells[Row, 110] = "Folio";
                worksheet.Cells[Row, 111] = "Nombre del agente";
                worksheet.Cells[Row, 112] = "Razon social";
                worksheet.Cells[Row, 113] = "Fecha de vencimiento";
                worksheet.Cells[Row, 114] = "RFC";
                worksheet.Cells[Row, 115] = "Subtotal";
                worksheet.Cells[Row, 116] = "IVA";
                worksheet.Cells[Row, 117] = "TOTAL";
                worksheet.Cells[Row, 118] = "Pendiente";
                worksheet.Cells[Row, 119] = "Texto Extra 3";
                worksheet.Cells[Row, 120] = "Afectado";
                worksheet.Cells[Row, 121] = "Impreso";
                worksheet.Cells[Row, 122] = "Cancelado";
                worksheet.Cells[Row, 123] = "Total de unidades";
                worksheet.Cells[Row, 124] = "Clasificacion cliente2";
                worksheet.Cells[Row, 125] = "Zona";
                worksheet.Cells[Row, 126] = "Agente";

                int column_mora = 21;

                //titulo 
                worksheet.Cells[2, 112 + column_mora] = "Abonos los 3 agentes";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 108 + column_mora] = "Fecha";
                worksheet.Cells[Row, 109 + column_mora] = "Serie";
                worksheet.Cells[Row, 110 + column_mora] = "Folio";
                worksheet.Cells[Row, 111 + column_mora] = "Nombre del agente";
                worksheet.Cells[Row, 112 + column_mora] = "Razon social";
                worksheet.Cells[Row, 113 + column_mora] = "Fecha de vencimiento";
                worksheet.Cells[Row, 114 + column_mora] = "RFC";
                worksheet.Cells[Row, 115 + column_mora] = "Subtotal";
                worksheet.Cells[Row, 116 + column_mora] = "IVA";
                worksheet.Cells[Row, 117 + column_mora] = "TOTAL";
                worksheet.Cells[Row, 118 + column_mora] = "Pendiente";
                worksheet.Cells[Row, 119 + column_mora] = "Texto Extra 3";
                worksheet.Cells[Row, 120 + column_mora] = "Afectado";
                worksheet.Cells[Row, 121 + column_mora] = "Impreso";
                worksheet.Cells[Row, 122 + column_mora] = "Cancelado";
                worksheet.Cells[Row, 123 + column_mora] = "Total de unidades";
                worksheet.Cells[Row, 124 + column_mora] = "Clasificacion cliente2";
                worksheet.Cells[Row, 125 + column_mora] = "Zona";
                worksheet.Cells[Row, 126 + column_mora] = "Agente";

                worksheet.Cells[2, 133 + column_mora] = "Zonas centro";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 129 + column_mora] = "Fecha";
                worksheet.Cells[Row, 130 + column_mora] = "Serie";
                worksheet.Cells[Row, 131 + column_mora] = "Folio";
                worksheet.Cells[Row, 132 + column_mora] = "Nombre del agente";
                worksheet.Cells[Row, 133 + column_mora] = "Razon social";
                worksheet.Cells[Row, 134 + column_mora] = "Fecha de vencimiento";
                worksheet.Cells[Row, 135 + column_mora] = "RFC";
                worksheet.Cells[Row, 136 + column_mora] = "Subtotal";
                worksheet.Cells[Row, 137 + column_mora] = "IVA";
                worksheet.Cells[Row, 138 + column_mora] = "TOTAL";
                worksheet.Cells[Row, 139 + column_mora] = "Pendiente";
                worksheet.Cells[Row, 140 + column_mora] = "Texto Extra 3";
                worksheet.Cells[Row, 141 + column_mora] = "Afectado";
                worksheet.Cells[Row, 142 + column_mora] = "Impreso";
                worksheet.Cells[Row, 143 + column_mora] = "Cancelado";
                worksheet.Cells[Row, 144 + column_mora] = "Total de unidades";
                worksheet.Cells[Row, 145 + column_mora] = "Zona ";
                worksheet.Cells[Row, 146 + column_mora] = "Agente";

                worksheet.Cells[2, 153 + column_mora] = "Zonas sur";

                worksheet.Cells[Row, 149 + column_mora] = "Fecha";
                worksheet.Cells[Row, 150 + column_mora] = "Serie";
                worksheet.Cells[Row, 151 + column_mora] = "Folio";
                worksheet.Cells[Row, 152 + column_mora] = "Nombre del agente";
                worksheet.Cells[Row, 153 + column_mora] = "Razon social";
                worksheet.Cells[Row, 154 + column_mora] = "Fecha de vencimiento";
                worksheet.Cells[Row, 155 + column_mora] = "RFC";
                worksheet.Cells[Row, 156 + column_mora] = "Subtotal";
                worksheet.Cells[Row, 157 + column_mora] = "IVA";
                worksheet.Cells[Row, 158 + column_mora] = "TOTAL";
                worksheet.Cells[Row, 159 + column_mora] = "Pendiente";
                worksheet.Cells[Row, 160 + column_mora] = "Texto Extra 3";
                worksheet.Cells[Row, 161 + column_mora] = "Afectado";
                worksheet.Cells[Row, 162 + column_mora] = "Impreso";
                worksheet.Cells[Row, 163 + column_mora] = "Cancelado";
                worksheet.Cells[Row, 164 + column_mora] = "Total de unidades";
                worksheet.Cells[Row, 165 + column_mora] = "Zona";
                worksheet.Cells[Row, 166 + column_mora] = "Agente";

                worksheet.Cells[2, 173 + column_mora] = "Zonas norte";

                worksheet.Cells[Row, 169 + column_mora] = "Fecha";
                worksheet.Cells[Row, 170 + column_mora] = "Serie";
                worksheet.Cells[Row, 171 + column_mora] = "Folio";
                worksheet.Cells[Row, 172 + column_mora] = "Nombre del agente";
                worksheet.Cells[Row, 173 + column_mora] = "Razon social";
                worksheet.Cells[Row, 174 + column_mora] = "Fecha de vencimiento";
                worksheet.Cells[Row, 175 + column_mora] = "RFC";
                worksheet.Cells[Row, 176 + column_mora] = "Subtotal";
                worksheet.Cells[Row, 177 + column_mora] = "IVA";
                worksheet.Cells[Row, 178 + column_mora] = "TOTAL";
                worksheet.Cells[Row, 179 + column_mora] = "Pendiente";
                worksheet.Cells[Row, 180 + column_mora] = "Texto Extra 3";
                worksheet.Cells[Row, 181 + column_mora] = "Afectado";
                worksheet.Cells[Row, 182 + column_mora] = "Impreso";
                worksheet.Cells[Row, 183 + column_mora] = "Cancelado";
                worksheet.Cells[Row, 184 + column_mora] = "Total de unidades";
                worksheet.Cells[Row, 185 + column_mora] = "Zona";
                worksheet.Cells[Row, 186 + column_mora] = "Agente";


                worksheet.Cells[2, 194 + column_mora] = "Cuarta Zona";

                worksheet.Cells[Row, 190 + column_mora] = "Fecha";
                worksheet.Cells[Row, 191 + column_mora] = "Serie";
                worksheet.Cells[Row, 192 + column_mora] = "Folio";
                worksheet.Cells[Row, 193 + column_mora] = "Nombre del agente";
                worksheet.Cells[Row, 194 + column_mora] = "Razon social";
                worksheet.Cells[Row, 195 + column_mora] = "Fecha de vencimiento";
                worksheet.Cells[Row, 196 + column_mora] = "RFC";
                worksheet.Cells[Row, 197 + column_mora] = "Subtotal";
                worksheet.Cells[Row, 198 + column_mora] = "IVA";
                worksheet.Cells[Row, 199 + column_mora] = "TOTAL";
                worksheet.Cells[Row, 200 + column_mora] = "Pendiente";
                worksheet.Cells[Row, 201 + column_mora] = "Texto Extra 3";
                worksheet.Cells[Row, 202 + column_mora] = "Afectado";
                worksheet.Cells[Row, 203 + column_mora] = "Impreso";
                worksheet.Cells[Row, 204 + column_mora] = "Cancelado";
                worksheet.Cells[Row, 205 + column_mora] = "Total de unidades";
                worksheet.Cells[Row, 206 + column_mora] = "Zona";
                worksheet.Cells[Row, 207 + column_mora] = "Agente";
                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.abonos.Count; i++)
                {
                    worksheet.Cells[Row, 67] = ListDocmuentos.abonos[i].Fecha;
                    worksheet.Cells[Row, 68] = ListDocmuentos.abonos[i].Serie;
                    worksheet.Cells[Row, 69] = ListDocmuentos.abonos[i].Folio;
                    worksheet.Cells[Row, 70] = ListDocmuentos.abonos[i].NombreAgente;
                    worksheet.Cells[Row, 71] = ListDocmuentos.abonos[i].RazonSocial;
                    worksheet.Cells[Row, 72] = ListDocmuentos.abonos[i].FechaVencimiento;
                    worksheet.Cells[Row, 73] = ListDocmuentos.abonos[i].RFC;
                    worksheet.Cells[Row, 74] = ListDocmuentos.abonos[i].Subtotal;
                    worksheet.Cells[Row, 75] = ListDocmuentos.abonos[i].IVA;
                    worksheet.Cells[Row, 76] = ListDocmuentos.abonos[i].Total;
                    worksheet.Cells[Row, 77] = ListDocmuentos.abonos[i].Pendiente;
                    worksheet.Cells[Row, 78] = ListDocmuentos.abonos[i].TextoExtra3;
                    worksheet.Cells[Row, 79] = ListDocmuentos.abonos[i].Afectado;
                    worksheet.Cells[Row, 80] = ListDocmuentos.abonos[i].Impreso;
                    worksheet.Cells[Row, 81] = ListDocmuentos.abonos[i].Cancelado;
                    worksheet.Cells[Row, 82] = ListDocmuentos.abonos[i].TotalUnidades;
                    worksheet.Cells[Row, 83] = ListDocmuentos.abonos[i].Clasificacion2;
                    worksheet.Cells[Row, 84] = ListDocmuentos.abonos[i].TextoExtra1;
                    worksheet.Cells[Row, 85] = ListDocmuentos.abonos[i].NombreConcepto;
                    total += ListDocmuentos.abonos[i].Total;
                    Row++;
                }
                //worksheet.Cells[2, 76] = "$ " + total;

                //

                for (int i = 0; i < ListDocmuentos.prestamos.Count; i++)
                {
                    worksheet.Cells[Row, 67] = ListDocmuentos.prestamos[i].Fecha;
                    worksheet.Cells[Row, 68] = ListDocmuentos.prestamos[i].Serie;
                    worksheet.Cells[Row, 69] = ListDocmuentos.prestamos[i].Folio;
                    worksheet.Cells[Row, 70] = ListDocmuentos.prestamos[i].NombreAgente;
                    worksheet.Cells[Row, 71] = ListDocmuentos.prestamos[i].RazonSocial;
                    worksheet.Cells[Row, 72] = ListDocmuentos.prestamos[i].FechaVencimiento;
                    worksheet.Cells[Row, 73] = ListDocmuentos.prestamos[i].RFC;
                    worksheet.Cells[Row, 74] = ListDocmuentos.prestamos[i].Subtotal;
                    worksheet.Cells[Row, 75] = ListDocmuentos.prestamos[i].IVA;
                    worksheet.Cells[Row, 76] = ListDocmuentos.prestamos[i].Total;
                    worksheet.Cells[Row, 77] = ListDocmuentos.prestamos[i].Pendiente;
                    worksheet.Cells[Row, 78] = ListDocmuentos.prestamos[i].TextoExtra3;
                    worksheet.Cells[Row, 79] = ListDocmuentos.prestamos[i].Afectado;
                    worksheet.Cells[Row, 80] = ListDocmuentos.prestamos[i].Impreso;
                    worksheet.Cells[Row, 81] = ListDocmuentos.prestamos[i].Cancelado;
                    worksheet.Cells[Row, 82] = ListDocmuentos.prestamos[i].TotalUnidades;
                    worksheet.Cells[Row, 83] = ListDocmuentos.prestamos[i].Clasificacion2;
                    worksheet.Cells[Row, 84] = ListDocmuentos.prestamos[i].TextoExtra1;
                    worksheet.Cells[Row, 85] = ListDocmuentos.prestamos[i].NombreConcepto;
                    total += ListDocmuentos.prestamos[i].Total;
                    Row++;
                }
                for (int i = 0; i < ListDocmuentos.ingreso_dev_garantia.Count; i++)
                {
                    worksheet.Cells[Row, 67] = ListDocmuentos.ingreso_dev_garantia[i].Fecha;
                    worksheet.Cells[Row, 68] = ListDocmuentos.ingreso_dev_garantia[i].Serie;
                    worksheet.Cells[Row, 69] = ListDocmuentos.ingreso_dev_garantia[i].Folio;
                    worksheet.Cells[Row, 70] = ListDocmuentos.ingreso_dev_garantia[i].NombreAgente;
                    worksheet.Cells[Row, 71] = ListDocmuentos.ingreso_dev_garantia[i].RazonSocial;
                    worksheet.Cells[Row, 72] = ListDocmuentos.ingreso_dev_garantia[i].FechaVencimiento;
                    worksheet.Cells[Row, 73] = ListDocmuentos.ingreso_dev_garantia[i].RFC;
                    worksheet.Cells[Row, 74] = ListDocmuentos.ingreso_dev_garantia[i].Subtotal;
                    worksheet.Cells[Row, 75] = ListDocmuentos.ingreso_dev_garantia[i].IVA;
                    worksheet.Cells[Row, 76] = ListDocmuentos.ingreso_dev_garantia[i].Total;
                    worksheet.Cells[Row, 77] = ListDocmuentos.ingreso_dev_garantia[i].Pendiente;
                    worksheet.Cells[Row, 78] = ListDocmuentos.ingreso_dev_garantia[i].TextoExtra3;
                    worksheet.Cells[Row, 79] = ListDocmuentos.ingreso_dev_garantia[i].Afectado;
                    worksheet.Cells[Row, 80] = ListDocmuentos.ingreso_dev_garantia[i].Impreso;
                    worksheet.Cells[Row, 81] = ListDocmuentos.ingreso_dev_garantia[i].Cancelado;
                    worksheet.Cells[Row, 82] = ListDocmuentos.ingreso_dev_garantia[i].TotalUnidades;
                    worksheet.Cells[Row, 83] = ListDocmuentos.ingreso_dev_garantia[i].Clasificacion2;
                    worksheet.Cells[Row, 84] = ListDocmuentos.ingreso_dev_garantia[i].TextoExtra1;
                    worksheet.Cells[Row, 85] = ListDocmuentos.ingreso_dev_garantia[i].NombreConcepto;
                    total += ListDocmuentos.ingreso_dev_garantia[i].Total;
                    Row++;
                }
                for (int i = 0; i < ListDocmuentos.ingreso_traspaso.Count; i++)
                {
                    worksheet.Cells[Row, 67] = ListDocmuentos.ingreso_traspaso[i].Fecha;
                    worksheet.Cells[Row, 68] = ListDocmuentos.ingreso_traspaso[i].Serie;
                    worksheet.Cells[Row, 69] = ListDocmuentos.ingreso_traspaso[i].Folio;
                    worksheet.Cells[Row, 70] = ListDocmuentos.ingreso_traspaso[i].NombreAgente;
                    worksheet.Cells[Row, 71] = ListDocmuentos.ingreso_traspaso[i].RazonSocial;
                    worksheet.Cells[Row, 72] = ListDocmuentos.ingreso_traspaso[i].FechaVencimiento;
                    worksheet.Cells[Row, 73] = ListDocmuentos.ingreso_traspaso[i].RFC;
                    worksheet.Cells[Row, 74] = ListDocmuentos.ingreso_traspaso[i].Subtotal;
                    worksheet.Cells[Row, 75] = ListDocmuentos.ingreso_traspaso[i].IVA;
                    worksheet.Cells[Row, 76] = ListDocmuentos.ingreso_traspaso[i].Total;
                    worksheet.Cells[Row, 77] = ListDocmuentos.ingreso_traspaso[i].Pendiente;
                    worksheet.Cells[Row, 78] = ListDocmuentos.ingreso_traspaso[i].TextoExtra3;
                    worksheet.Cells[Row, 79] = ListDocmuentos.ingreso_traspaso[i].Afectado;
                    worksheet.Cells[Row, 80] = ListDocmuentos.ingreso_traspaso[i].Impreso;
                    worksheet.Cells[Row, 81] = ListDocmuentos.ingreso_traspaso[i].Cancelado;
                    worksheet.Cells[Row, 82] = ListDocmuentos.ingreso_traspaso[i].TotalUnidades;
                    worksheet.Cells[Row, 83] = ListDocmuentos.ingreso_traspaso[i].Clasificacion2;
                    worksheet.Cells[Row, 84] = ListDocmuentos.ingreso_traspaso[i].TextoExtra1;
                    worksheet.Cells[Row, 85] = ListDocmuentos.ingreso_traspaso[i].NombreConcepto;
                    total += ListDocmuentos.ingreso_traspaso[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 76] = "$ " + total;

                //
                total = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.abonos_rfc_publico.Count; i++)
                {
                    worksheet.Cells[Row, 88] = ListDocmuentos.abonos_rfc_publico[i].Fecha;
                    worksheet.Cells[Row, 89] = ListDocmuentos.abonos_rfc_publico[i].Serie;
                    worksheet.Cells[Row, 90] = ListDocmuentos.abonos_rfc_publico[i].Folio;
                    worksheet.Cells[Row, 91] = ListDocmuentos.abonos_rfc_publico[i].NombreAgente;
                    worksheet.Cells[Row, 92] = ListDocmuentos.abonos_rfc_publico[i].RazonSocial;
                    worksheet.Cells[Row, 93] = ListDocmuentos.abonos_rfc_publico[i].FechaVencimiento;
                    worksheet.Cells[Row, 94] = ListDocmuentos.abonos_rfc_publico[i].RFC;
                    worksheet.Cells[Row, 95] = ListDocmuentos.abonos_rfc_publico[i].Subtotal;
                    worksheet.Cells[Row, 96] = ListDocmuentos.abonos_rfc_publico[i].IVA;
                    worksheet.Cells[Row, 97] = ListDocmuentos.abonos_rfc_publico[i].Total;
                    worksheet.Cells[Row, 98] = ListDocmuentos.abonos_rfc_publico[i].Pendiente;
                    worksheet.Cells[Row, 99] = ListDocmuentos.abonos_rfc_publico[i].TextoExtra3;
                    worksheet.Cells[Row, 100] = ListDocmuentos.abonos_rfc_publico[i].Afectado;
                    worksheet.Cells[Row, 101] = ListDocmuentos.abonos_rfc_publico[i].Impreso;
                    worksheet.Cells[Row, 102] = ListDocmuentos.abonos_rfc_publico[i].Cancelado;
                    worksheet.Cells[Row, 103] = ListDocmuentos.abonos_rfc_publico[i].TotalUnidades;
                    worksheet.Cells[Row, 104] = ListDocmuentos.abonos_rfc_publico[i].Clasificacion2;
                    worksheet.Cells[Row, 105] = ListDocmuentos.abonos_rfc_publico[i].TextoExtra1;
                    worksheet.Cells[Row, 106] = ListDocmuentos.abonos_rfc_publico[i].NombreConcepto;
                    total += ListDocmuentos.abonos_rfc_publico[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 97] = "$ " + total;
                total = 0;
                Row = 5;
                for (int i = 0; i < ListDocmuentos.abonos_ol.Count; i++)
                {
                    worksheet.Cells[Row, 108] = ListDocmuentos.abonos_ol[i].Fecha;
                    worksheet.Cells[Row, 109] = ListDocmuentos.abonos_ol[i].Serie;
                    worksheet.Cells[Row, 110] = ListDocmuentos.abonos_ol[i].Folio;
                    worksheet.Cells[Row, 111] = ListDocmuentos.abonos_ol[i].NombreAgente;
                    worksheet.Cells[Row, 112] = ListDocmuentos.abonos_ol[i].RazonSocial;
                    worksheet.Cells[Row, 113] = ListDocmuentos.abonos_ol[i].FechaVencimiento;
                    worksheet.Cells[Row, 114] = ListDocmuentos.abonos_ol[i].RFC;
                    worksheet.Cells[Row, 115] = ListDocmuentos.abonos_ol[i].Subtotal;
                    worksheet.Cells[Row, 116] = ListDocmuentos.abonos_ol[i].IVA;
                    worksheet.Cells[Row, 117] = ListDocmuentos.abonos_ol[i].Total;
                    worksheet.Cells[Row, 118] = ListDocmuentos.abonos_ol[i].Pendiente;
                    worksheet.Cells[Row, 119] = ListDocmuentos.abonos_ol[i].TextoExtra3;
                    worksheet.Cells[Row, 120] = ListDocmuentos.abonos_ol[i].Afectado;
                    worksheet.Cells[Row, 121] = ListDocmuentos.abonos_ol[i].Impreso;
                    worksheet.Cells[Row, 122] = ListDocmuentos.abonos_ol[i].Cancelado;
                    worksheet.Cells[Row, 123] = ListDocmuentos.abonos_ol[i].TotalUnidades;
                    worksheet.Cells[Row, 124] = ListDocmuentos.abonos_ol[i].Clasificacion2;
                    worksheet.Cells[Row, 125] = ListDocmuentos.abonos_ol[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 126] = ListDocmuentos.abonos_ol[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_ol[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 117] = "$ " + total;



                // + column_mora
                total = 0;
                Row = 5;
                for (int i = 0; i < ListDocmuentos.abonos_zona_centro.Count; i++)
                {
                    worksheet.Cells[Row, 108 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Fecha;
                    worksheet.Cells[Row, 109 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Serie;
                    worksheet.Cells[Row, 110 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Folio;
                    worksheet.Cells[Row, 111 + column_mora] = ListDocmuentos.abonos_zona_centro[i].NombreAgente;
                    worksheet.Cells[Row, 112 + column_mora] = ListDocmuentos.abonos_zona_centro[i].RazonSocial;
                    worksheet.Cells[Row, 113 + column_mora] = ListDocmuentos.abonos_zona_centro[i].FechaVencimiento;
                    worksheet.Cells[Row, 114 + column_mora] = ListDocmuentos.abonos_zona_centro[i].RFC;
                    worksheet.Cells[Row, 115 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Subtotal;
                    worksheet.Cells[Row, 116 + column_mora] = ListDocmuentos.abonos_zona_centro[i].IVA;
                    worksheet.Cells[Row, 117 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Total;
                    worksheet.Cells[Row, 118 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Pendiente;
                    worksheet.Cells[Row, 119 + column_mora] = ListDocmuentos.abonos_zona_centro[i].TextoExtra3;
                    worksheet.Cells[Row, 120 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Afectado;
                    worksheet.Cells[Row, 121 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Impreso;
                    worksheet.Cells[Row, 122 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Cancelado;
                    worksheet.Cells[Row, 123 + column_mora] = ListDocmuentos.abonos_zona_centro[i].TotalUnidades;
                    worksheet.Cells[Row, 124 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Clasificacion2;
                    worksheet.Cells[Row, 125 + column_mora] = ListDocmuentos.abonos_zona_centro[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 126 + column_mora] = ListDocmuentos.abonos_zona_centro[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_centro[i].Total;
                    Row++;
                }
                for (int i = 0; i < ListDocmuentos.abonos_zona_norte.Count; i++)
                {
                    worksheet.Cells[Row, 108 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Fecha;
                    worksheet.Cells[Row, 109 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Serie;
                    worksheet.Cells[Row, 110 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Folio;
                    worksheet.Cells[Row, 111 + column_mora] = ListDocmuentos.abonos_zona_norte[i].NombreAgente;
                    worksheet.Cells[Row, 112 + column_mora] = ListDocmuentos.abonos_zona_norte[i].RazonSocial;
                    worksheet.Cells[Row, 113 + column_mora] = ListDocmuentos.abonos_zona_norte[i].FechaVencimiento;
                    worksheet.Cells[Row, 114 + column_mora] = ListDocmuentos.abonos_zona_norte[i].RFC;
                    worksheet.Cells[Row, 115 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Subtotal;
                    worksheet.Cells[Row, 116 + column_mora] = ListDocmuentos.abonos_zona_norte[i].IVA;
                    worksheet.Cells[Row, 117 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Total;
                    worksheet.Cells[Row, 118 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Pendiente;
                    worksheet.Cells[Row, 119 + column_mora] = ListDocmuentos.abonos_zona_norte[i].TextoExtra3;
                    worksheet.Cells[Row, 120 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Afectado;
                    worksheet.Cells[Row, 121 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Impreso;
                    worksheet.Cells[Row, 122 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Cancelado;
                    worksheet.Cells[Row, 123 + column_mora] = ListDocmuentos.abonos_zona_norte[i].TotalUnidades;
                    worksheet.Cells[Row, 124 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Clasificacion2;
                    worksheet.Cells[Row, 125 + column_mora] = ListDocmuentos.abonos_zona_norte[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 126 + column_mora] = ListDocmuentos.abonos_zona_norte[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_norte[i].Total;
                    Row++;
                }
                for (int i = 0; i < ListDocmuentos.abonos_zona_sur.Count; i++)
                {
                    worksheet.Cells[Row, 108 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Fecha;
                    worksheet.Cells[Row, 109 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Serie;
                    worksheet.Cells[Row, 110 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Folio;
                    worksheet.Cells[Row, 111 + column_mora] = ListDocmuentos.abonos_zona_sur[i].NombreAgente;
                    worksheet.Cells[Row, 112 + column_mora] = ListDocmuentos.abonos_zona_sur[i].RazonSocial;
                    worksheet.Cells[Row, 113 + column_mora] = ListDocmuentos.abonos_zona_sur[i].FechaVencimiento;
                    worksheet.Cells[Row, 114 + column_mora] = ListDocmuentos.abonos_zona_sur[i].RFC;
                    worksheet.Cells[Row, 115 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Subtotal;
                    worksheet.Cells[Row, 116 + column_mora] = ListDocmuentos.abonos_zona_sur[i].IVA;
                    worksheet.Cells[Row, 117 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Total;
                    worksheet.Cells[Row, 118 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Pendiente;
                    worksheet.Cells[Row, 119 + column_mora] = ListDocmuentos.abonos_zona_sur[i].TextoExtra3;
                    worksheet.Cells[Row, 120 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Afectado;
                    worksheet.Cells[Row, 121 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Impreso;
                    worksheet.Cells[Row, 122 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Cancelado;
                    worksheet.Cells[Row, 123 + column_mora] = ListDocmuentos.abonos_zona_sur[i].TotalUnidades;
                    worksheet.Cells[Row, 124 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Clasificacion2;
                    worksheet.Cells[Row, 125 + column_mora] = ListDocmuentos.abonos_zona_sur[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 126 + column_mora] = ListDocmuentos.abonos_zona_sur[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_sur[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 117 + column_mora] = "$ " + total;
                total = 0;
                Row = 5;


                /*ZONA CENTRO*/
                for (int i = 0; i < ListDocmuentos.abonos_zona_centro.Count; i++)
                {
                    worksheet.Cells[Row, 129 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Fecha;
                    worksheet.Cells[Row, 130 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Serie;
                    worksheet.Cells[Row, 131 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Folio;
                    worksheet.Cells[Row, 132 + column_mora] = ListDocmuentos.abonos_zona_centro[i].NombreAgente;
                    worksheet.Cells[Row, 133 + column_mora] = ListDocmuentos.abonos_zona_centro[i].RazonSocial;
                    worksheet.Cells[Row, 134 + column_mora] = ListDocmuentos.abonos_zona_centro[i].FechaVencimiento;
                    worksheet.Cells[Row, 135 + column_mora] = ListDocmuentos.abonos_zona_centro[i].RFC;
                    worksheet.Cells[Row, 136 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Subtotal;
                    worksheet.Cells[Row, 137 + column_mora] = ListDocmuentos.abonos_zona_centro[i].IVA;
                    worksheet.Cells[Row, 138 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Total;
                    worksheet.Cells[Row, 139 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Pendiente;
                    worksheet.Cells[Row, 140 + column_mora] = ListDocmuentos.abonos_zona_centro[i].TextoExtra3;
                    worksheet.Cells[Row, 141 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Afectado;
                    worksheet.Cells[Row, 142 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Impreso;
                    worksheet.Cells[Row, 143 + column_mora] = ListDocmuentos.abonos_zona_centro[i].Cancelado;
                    worksheet.Cells[Row, 144 + column_mora] = ListDocmuentos.abonos_zona_centro[i].TotalUnidades;
                    worksheet.Cells[Row, 145 + column_mora] = ListDocmuentos.abonos_zona_centro[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 146 + column_mora] = ListDocmuentos.abonos_zona_centro[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_centro[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 138 + column_mora] = "$ " + total;
                total = 0;
                Row = 5;
                /*ZONA sur*/
                for (int i = 0; i < ListDocmuentos.abonos_zona_sur.Count; i++)
                {
                    worksheet.Cells[Row, 149 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Fecha;
                    worksheet.Cells[Row, 150 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Serie;
                    worksheet.Cells[Row, 151 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Folio;
                    worksheet.Cells[Row, 152 + column_mora] = ListDocmuentos.abonos_zona_sur[i].NombreAgente;
                    worksheet.Cells[Row, 153 + column_mora] = ListDocmuentos.abonos_zona_sur[i].RazonSocial;
                    worksheet.Cells[Row, 154 + column_mora] = ListDocmuentos.abonos_zona_sur[i].FechaVencimiento;
                    worksheet.Cells[Row, 155 + column_mora] = ListDocmuentos.abonos_zona_sur[i].RFC;
                    worksheet.Cells[Row, 156 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Subtotal;
                    worksheet.Cells[Row, 157 + column_mora] = ListDocmuentos.abonos_zona_sur[i].IVA;
                    worksheet.Cells[Row, 158 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Total;
                    worksheet.Cells[Row, 159 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Pendiente;
                    worksheet.Cells[Row, 160 + column_mora] = ListDocmuentos.abonos_zona_sur[i].TextoExtra3;
                    worksheet.Cells[Row, 161 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Afectado;
                    worksheet.Cells[Row, 162 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Impreso;
                    worksheet.Cells[Row, 163 + column_mora] = ListDocmuentos.abonos_zona_sur[i].Cancelado;
                    worksheet.Cells[Row, 164 + column_mora] = ListDocmuentos.abonos_zona_sur[i].TotalUnidades;
                    worksheet.Cells[Row, 165 + column_mora] = ListDocmuentos.abonos_zona_sur[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 166 + column_mora] = ListDocmuentos.abonos_zona_sur[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_sur[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 158 + column_mora] = "$ " + total;
                total = 0;
                Row = 5;
                /*ZONA norte*/
                for (int i = 0; i < ListDocmuentos.abonos_zona_norte.Count; i++)
                {
                    worksheet.Cells[Row, 169 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Fecha;
                    worksheet.Cells[Row, 170 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Serie;
                    worksheet.Cells[Row, 171 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Folio;
                    worksheet.Cells[Row, 172 + column_mora] = ListDocmuentos.abonos_zona_norte[i].NombreAgente;
                    worksheet.Cells[Row, 173 + column_mora] = ListDocmuentos.abonos_zona_norte[i].RazonSocial;
                    worksheet.Cells[Row, 174 + column_mora] = ListDocmuentos.abonos_zona_norte[i].FechaVencimiento;
                    worksheet.Cells[Row, 175 + column_mora] = ListDocmuentos.abonos_zona_norte[i].RFC;
                    worksheet.Cells[Row, 176 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Subtotal;
                    worksheet.Cells[Row, 177 + column_mora] = ListDocmuentos.abonos_zona_norte[i].IVA;
                    worksheet.Cells[Row, 178 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Total;
                    worksheet.Cells[Row, 179 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Pendiente;
                    worksheet.Cells[Row, 180 + column_mora] = ListDocmuentos.abonos_zona_norte[i].TextoExtra3;
                    worksheet.Cells[Row, 181 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Afectado;
                    worksheet.Cells[Row, 182 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Impreso;
                    worksheet.Cells[Row, 183 + column_mora] = ListDocmuentos.abonos_zona_norte[i].Cancelado;
                    worksheet.Cells[Row, 184 + column_mora] = ListDocmuentos.abonos_zona_norte[i].TotalUnidades;
                    worksheet.Cells[Row, 185 + column_mora] = ListDocmuentos.abonos_zona_norte[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 186 + column_mora] = ListDocmuentos.abonos_zona_norte[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_norte[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 178 + column_mora] = "$ " + total;

                total = 0;
                Row = 5;
                /*ZONA norte*/
                for (int i = 0; i < ListDocmuentos.abonos_zona_cuatro.Count; i++)
                {
                    worksheet.Cells[Row, 190 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Fecha;
                    worksheet.Cells[Row, 191 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Serie;
                    worksheet.Cells[Row, 192 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Folio;
                    worksheet.Cells[Row, 193 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].NombreAgente;
                    worksheet.Cells[Row, 194 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].RazonSocial;
                    worksheet.Cells[Row, 195 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].FechaVencimiento;
                    worksheet.Cells[Row, 196 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].RFC;
                    worksheet.Cells[Row, 197 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Subtotal;
                    worksheet.Cells[Row, 198 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].IVA;
                    worksheet.Cells[Row, 199 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Total;
                    worksheet.Cells[Row, 200 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Pendiente;
                    worksheet.Cells[Row, 201 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].TextoExtra3;
                    worksheet.Cells[Row, 202 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Afectado;
                    worksheet.Cells[Row, 203 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Impreso;
                    worksheet.Cells[Row, 204 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].Cancelado;
                    worksheet.Cells[Row, 205 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].TotalUnidades;
                    worksheet.Cells[Row, 206 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 207 + column_mora] = ListDocmuentos.abonos_zona_cuatro[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_cuatro[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 199 + column_mora] = "$ " + total;
                #endregion
                
                #endregion
                int new_row = 22 + column_mora;
                #region Compras
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 194 + new_row] = "compras acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 190 + new_row] = "Fecha";
                worksheet.Cells[Row, 191 + new_row] = "Serie";
                worksheet.Cells[Row, 192 + new_row] = "Folio";
                worksheet.Cells[Row, 193 + new_row] = "Nombre del agente";
                worksheet.Cells[Row, 194 + new_row] = "Razon social";
                worksheet.Cells[Row, 195 + new_row] = "Fecha de vencimiento";
                worksheet.Cells[Row, 196 + new_row] = "RFC";
                worksheet.Cells[Row, 197 + new_row] = "Subtotal";
                worksheet.Cells[Row, 198 + new_row] = "IVA";
                worksheet.Cells[Row, 199 + new_row] = "TOTAL";
                worksheet.Cells[Row, 200 + new_row] = "Pendiente";
                worksheet.Cells[Row, 201 + new_row] = "Texto Extra 3";
                worksheet.Cells[Row, 202 + new_row] = "Afectado";
                worksheet.Cells[Row, 203 + new_row] = "Impreso";
                worksheet.Cells[Row, 204 + new_row] = "Cancelado";
                worksheet.Cells[Row, 205 + new_row] = "Total de unidades";
                worksheet.Cells[Row, 206 + new_row] = "Clasificacion cliente2";
                worksheet.Cells[Row, 207 + new_row] = "Texto extra1";
                worksheet.Cells[Row, 208 + new_row] = "Nombre del concepto";
                //titulo 
                worksheet.Cells[2, 215 + new_row] = "compras ANJI";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 211 + new_row] = "Fecha";
                worksheet.Cells[Row, 212 + new_row] = "Serie";
                worksheet.Cells[Row, 213 + new_row] = "Folio";
                worksheet.Cells[Row, 214 + new_row] = "Nombre del agente";
                worksheet.Cells[Row, 215 + new_row] = "Razon social";
                worksheet.Cells[Row, 216 + new_row] = "Fecha de vencimiento";
                worksheet.Cells[Row, 217 + new_row] = "RFC";
                worksheet.Cells[Row, 218 + new_row] = "Subtotal";
                worksheet.Cells[Row, 219 + new_row] = "IVA";
                worksheet.Cells[Row, 220 + new_row] = "TOTAL";
                worksheet.Cells[Row, 221 + new_row] = "Pendiente";
                worksheet.Cells[Row, 222 + new_row] = "Texto Extra 3";
                worksheet.Cells[Row, 223 + new_row] = "Afectado";
                worksheet.Cells[Row, 224 + new_row] = "Impreso";
                worksheet.Cells[Row, 225 + new_row] = "Cancelado";
                worksheet.Cells[Row, 226 + new_row] = "Total de unidades";
                worksheet.Cells[Row, 227 + new_row] = "Clasificacion cliente2";
                worksheet.Cells[Row, 228 + new_row] = "Texto extra1";
                worksheet.Cells[Row, 229 + new_row] = "Nombre del concepto";

                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.compras.Count; i++)
                {
                    worksheet.Cells[Row, 190 + new_row] = ListDocmuentos.compras[i].Fecha;
                    worksheet.Cells[Row, 191 + new_row] = ListDocmuentos.compras[i].Serie;
                    worksheet.Cells[Row, 192 + new_row] = ListDocmuentos.compras[i].Folio;
                    worksheet.Cells[Row, 193 + new_row] = ListDocmuentos.compras[i].NombreAgente;
                    worksheet.Cells[Row, 194 + new_row] = ListDocmuentos.compras[i].RazonSocial;
                    worksheet.Cells[Row, 195 + new_row] = ListDocmuentos.compras[i].FechaVencimiento;
                    worksheet.Cells[Row, 196 + new_row] = ListDocmuentos.compras[i].RFC;
                    worksheet.Cells[Row, 197 + new_row] = ListDocmuentos.compras[i].Subtotal;
                    worksheet.Cells[Row, 198 + new_row] = ListDocmuentos.compras[i].IVA;
                    worksheet.Cells[Row, 199 + new_row] = ListDocmuentos.compras[i].Total;
                    worksheet.Cells[Row, 200 + new_row] = ListDocmuentos.compras[i].Pendiente;
                    worksheet.Cells[Row, 201 + new_row] = ListDocmuentos.compras[i].TextoExtra3;
                    worksheet.Cells[Row, 202 + new_row] = ListDocmuentos.compras[i].Afectado;
                    worksheet.Cells[Row, 203 + new_row] = ListDocmuentos.compras[i].Impreso;
                    worksheet.Cells[Row, 204 + new_row] = ListDocmuentos.compras[i].Cancelado;
                    worksheet.Cells[Row, 205 + new_row] = ListDocmuentos.compras[i].TotalUnidades;
                    worksheet.Cells[Row, 206 + new_row] = ListDocmuentos.compras[i].Clasificacion2;
                    worksheet.Cells[Row, 207 + new_row] = ListDocmuentos.compras[i].TextoExtra1;
                    worksheet.Cells[Row, 208 + new_row] = ListDocmuentos.compras[i].NombreConcepto;
                    total += ListDocmuentos.compras[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 199 + new_row] = "$ " + total;

                total = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.compras_rfc_anji.Count; i++)
                {
                    worksheet.Cells[Row, 211 + new_row] = ListDocmuentos.compras_rfc_anji[i].Fecha;
                    worksheet.Cells[Row, 212 + new_row] = ListDocmuentos.compras_rfc_anji[i].Serie;
                    worksheet.Cells[Row, 213 + new_row] = ListDocmuentos.compras_rfc_anji[i].Folio;
                    worksheet.Cells[Row, 214 + new_row] = ListDocmuentos.compras_rfc_anji[i].NombreAgente;
                    worksheet.Cells[Row, 215 + new_row] = ListDocmuentos.compras_rfc_anji[i].RazonSocial;
                    worksheet.Cells[Row, 216 + new_row] = ListDocmuentos.compras_rfc_anji[i].FechaVencimiento;
                    worksheet.Cells[Row, 217 + new_row] = ListDocmuentos.compras_rfc_anji[i].RFC;
                    worksheet.Cells[Row, 218 + new_row] = ListDocmuentos.compras_rfc_anji[i].Subtotal;
                    worksheet.Cells[Row, 219 + new_row] = ListDocmuentos.compras_rfc_anji[i].IVA;
                    worksheet.Cells[Row, 220 + new_row] = ListDocmuentos.compras_rfc_anji[i].Total;
                    worksheet.Cells[Row, 221 + new_row] = ListDocmuentos.compras_rfc_anji[i].Pendiente;
                    worksheet.Cells[Row, 222 + new_row] = ListDocmuentos.compras_rfc_anji[i].TextoExtra3;
                    worksheet.Cells[Row, 223 + new_row] = ListDocmuentos.compras_rfc_anji[i].Afectado;
                    worksheet.Cells[Row, 224 + new_row] = ListDocmuentos.compras_rfc_anji[i].Impreso;
                    worksheet.Cells[Row, 225 + new_row] = ListDocmuentos.compras_rfc_anji[i].Cancelado;
                    worksheet.Cells[Row, 226 + new_row] = ListDocmuentos.compras_rfc_anji[i].TotalUnidades;
                    worksheet.Cells[Row, 227 + new_row] = ListDocmuentos.compras_rfc_anji[i].Clasificacion2;
                    worksheet.Cells[Row, 228 + new_row] = ListDocmuentos.compras_rfc_anji[i].TextoExtra1;
                    worksheet.Cells[Row, 229 + new_row] = ListDocmuentos.compras_rfc_anji[i].NombreConcepto;
                    total += ListDocmuentos.compras_rfc_anji[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 220 + new_row] = "$ " + total;
                #endregion
                #endregion

                #region Pagos proveedor
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 238 + new_row] = "pagos acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 234 + new_row] = "Fecha";
                worksheet.Cells[Row, 235 + new_row] = "Serie";
                worksheet.Cells[Row, 236 + new_row] = "Folio";
                worksheet.Cells[Row, 237 + new_row] = "Nombre del agente";
                worksheet.Cells[Row, 238 + new_row] = "Razon social";
                worksheet.Cells[Row, 239 + new_row] = "Fecha de vencimiento";
                worksheet.Cells[Row, 240 + new_row] = "RFC";
                worksheet.Cells[Row, 241 + new_row] = "Subtotal";
                worksheet.Cells[Row, 242 + new_row] = "IVA";
                worksheet.Cells[Row, 243 + new_row] = "TOTAL";
                worksheet.Cells[Row, 244 + new_row] = "Pendiente";
                worksheet.Cells[Row, 245 + new_row] = "Texto Extra 3";
                worksheet.Cells[Row, 246 + new_row] = "Afectado";
                worksheet.Cells[Row, 247 + new_row] = "Impreso";
                worksheet.Cells[Row, 248 + new_row] = "Cancelado";
                worksheet.Cells[Row, 249 + new_row] = "Total de unidades";
                worksheet.Cells[Row, 250 + new_row] = "Clasificacion cliente2";
                worksheet.Cells[Row, 251 + new_row] = "Texto extra1";
                worksheet.Cells[Row, 252 + new_row] = "Nombre del concepto";
                //titulo 
                worksheet.Cells[2, 259 + new_row] = "pagos Anji";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 255 + new_row] = "Fecha";
                worksheet.Cells[Row, 256 + new_row] = "Serie";
                worksheet.Cells[Row, 257 + new_row] = "Folio";
                worksheet.Cells[Row, 258 + new_row] = "Nombre del agente";
                worksheet.Cells[Row, 259 + new_row] = "Razon social";
                worksheet.Cells[Row, 260 + new_row] = "Fecha de vencimiento";
                worksheet.Cells[Row, 261 + new_row] = "RFC";
                worksheet.Cells[Row, 262 + new_row] = "Subtotal";
                worksheet.Cells[Row, 263 + new_row] = "IVA";
                worksheet.Cells[Row, 264 + new_row] = "TOTAL";
                worksheet.Cells[Row, 265 + new_row] = "Pendiente";
                worksheet.Cells[Row, 266 + new_row] = "Texto Extra 3";
                worksheet.Cells[Row, 267 + new_row] = "Afectado";
                worksheet.Cells[Row, 268 + new_row] = "Impreso";
                worksheet.Cells[Row, 269 + new_row] = "Cancelado";
                worksheet.Cells[Row, 270 + new_row] = "Total de unidades";
                worksheet.Cells[Row, 271 + new_row] = "Clasificacion cliente2";
                worksheet.Cells[Row, 272 + new_row] = "Texto extra1";
                worksheet.Cells[Row, 273 + new_row] = "Nombre del concepto";

                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.pagos_proveedor.Count; i++)
                {
                    worksheet.Cells[Row, 234 + new_row] = ListDocmuentos.pagos_proveedor[i].Fecha;
                    worksheet.Cells[Row, 235 + new_row] = ListDocmuentos.pagos_proveedor[i].Serie;
                    worksheet.Cells[Row, 236 + new_row] = ListDocmuentos.pagos_proveedor[i].Folio;
                    worksheet.Cells[Row, 237 + new_row] = ListDocmuentos.pagos_proveedor[i].NombreAgente;
                    worksheet.Cells[Row, 238 + new_row] = ListDocmuentos.pagos_proveedor[i].RazonSocial;
                    worksheet.Cells[Row, 239 + new_row] = ListDocmuentos.pagos_proveedor[i].FechaVencimiento;
                    worksheet.Cells[Row, 240 + new_row] = ListDocmuentos.pagos_proveedor[i].RFC;
                    worksheet.Cells[Row, 241 + new_row] = ListDocmuentos.pagos_proveedor[i].Subtotal;
                    worksheet.Cells[Row, 242 + new_row] = ListDocmuentos.pagos_proveedor[i].IVA;
                    worksheet.Cells[Row, 243 + new_row] = ListDocmuentos.pagos_proveedor[i].Total;
                    worksheet.Cells[Row, 244 + new_row] = ListDocmuentos.pagos_proveedor[i].Pendiente;
                    worksheet.Cells[Row, 245 + new_row] = ListDocmuentos.pagos_proveedor[i].TextoExtra3;
                    worksheet.Cells[Row, 246 + new_row] = ListDocmuentos.pagos_proveedor[i].Afectado;
                    worksheet.Cells[Row, 247 + new_row] = ListDocmuentos.pagos_proveedor[i].Impreso;
                    worksheet.Cells[Row, 248 + new_row] = ListDocmuentos.pagos_proveedor[i].Cancelado;
                    worksheet.Cells[Row, 249 + new_row] = ListDocmuentos.pagos_proveedor[i].TotalUnidades;
                    worksheet.Cells[Row, 250 + new_row] = ListDocmuentos.pagos_proveedor[i].Clasificacion2;
                    worksheet.Cells[Row, 251 + new_row] = ListDocmuentos.pagos_proveedor[i].TextoExtra1;
                    worksheet.Cells[Row, 252 + new_row] = ListDocmuentos.pagos_proveedor[i].NombreConcepto;
                    total += ListDocmuentos.pagos_proveedor[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 243 + new_row] = "$ " + total;

                total = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.pagos_proveedor_rfc_anji.Count; i++)
                {
                    worksheet.Cells[Row, 255 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Fecha;
                    worksheet.Cells[Row, 256 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Serie;
                    worksheet.Cells[Row, 257 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Folio;
                    worksheet.Cells[Row, 258 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].NombreAgente;
                    worksheet.Cells[Row, 259 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].RazonSocial;
                    worksheet.Cells[Row, 260 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].FechaVencimiento;
                    worksheet.Cells[Row, 261 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].RFC;
                    worksheet.Cells[Row, 262 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Subtotal;
                    worksheet.Cells[Row, 263 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].IVA;
                    worksheet.Cells[Row, 264 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Total;
                    worksheet.Cells[Row, 265 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Pendiente;
                    worksheet.Cells[Row, 266 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].TextoExtra3;
                    worksheet.Cells[Row, 267 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Afectado;
                    worksheet.Cells[Row, 268 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Impreso;
                    worksheet.Cells[Row, 269 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Cancelado;
                    worksheet.Cells[Row, 270 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].TotalUnidades;
                    worksheet.Cells[Row, 271 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Clasificacion2;
                    worksheet.Cells[Row, 272 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].TextoExtra1;
                    worksheet.Cells[Row, 273 + new_row] = ListDocmuentos.pagos_proveedor_rfc_anji[i].NombreConcepto;
                    total += ListDocmuentos.pagos_proveedor_rfc_anji[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 264 + new_row] = "$ " + total;

                #endregion
                #endregion

                #region Prestamos
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 277 + new_row] = "Prestamos";

                //encabezados prestamos
                worksheet.Cells[Row, 277 + new_row] = "Fecha";
                worksheet.Cells[Row, 278 + new_row] = "Serie";
                worksheet.Cells[Row, 279 + new_row] = "Folio";
                worksheet.Cells[Row, 280 + new_row] = "Nombre del agente";
                worksheet.Cells[Row, 281 + new_row] = "Razon social";
                worksheet.Cells[Row, 282 + new_row] = "Fecha de vencimiento";
                worksheet.Cells[Row, 283 + new_row] = "TextoExtra1";
                worksheet.Cells[Row, 284 + new_row] = "RFC";
                worksheet.Cells[Row, 285 + new_row] = "TOTAL";
                worksheet.Cells[Row, 286 + new_row] = "Pendiente";
                worksheet.Cells[Row, 287 + new_row] = "Cuenta";
                worksheet.Cells[Row, 288 + new_row] = "Referencia";

                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.prestamos.Count; i++)
                {
                    worksheet.Cells[Row, 277 + new_row] = ListDocmuentos.prestamos[i].Fecha;
                    worksheet.Cells[Row, 278 + new_row] = ListDocmuentos.prestamos[i].Serie;
                    worksheet.Cells[Row, 279 + new_row] = ListDocmuentos.prestamos[i].Folio;
                    worksheet.Cells[Row, 280 + new_row] = ListDocmuentos.prestamos[i].NombreAgente;
                    worksheet.Cells[Row, 281 + new_row] = ListDocmuentos.prestamos[i].RazonSocial;
                    worksheet.Cells[Row, 282 + new_row] = ListDocmuentos.prestamos[i].FechaVencimiento;
                    worksheet.Cells[Row, 283 + new_row] = ListDocmuentos.prestamos[i].TextoExtra1;
                    worksheet.Cells[Row, 284 + new_row] = ListDocmuentos.prestamos[i].RFC;
                    worksheet.Cells[Row, 285 + new_row] = ListDocmuentos.prestamos[i].Total;
                    worksheet.Cells[Row, 286 + new_row] = ListDocmuentos.prestamos[i].Pendiente;
                    worksheet.Cells[Row, 287 + new_row] = ListDocmuentos.prestamos[i].TextoExtra2;
                    worksheet.Cells[Row, 288 + new_row] = ListDocmuentos.prestamos[i].Referencia;
                    total += ListDocmuentos.prestamos[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 285 + new_row] = "$ " + total;
                #endregion
                #endregion

                #region Ingreso traspaso
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 294 + new_row] = "Ingreso traspaso";

                //encabezados ingtreso traspaso
                worksheet.Cells[Row, 290 + new_row] = "Fecha";
                worksheet.Cells[Row, 291 + new_row] = "Serie";
                worksheet.Cells[Row, 292 + new_row] = "Folio";
                worksheet.Cells[Row, 293 + new_row] = "Nombre del agente";
                worksheet.Cells[Row, 294 + new_row] = "Razon social";
                worksheet.Cells[Row, 295 + new_row] = "Fecha de vencimiento";
                worksheet.Cells[Row, 296 + new_row] = "Fecha de depósito";
                worksheet.Cells[Row, 297 + new_row] = "RFC";
                worksheet.Cells[Row, 298 + new_row] = "TOTAL";
                worksheet.Cells[Row, 299 + new_row] = "Pendiente";
                worksheet.Cells[Row, 300 + new_row] = "texto extra 2";
                worksheet.Cells[Row, 301 + new_row] = "Referencia";

                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.ingreso_traspaso.Count; i++)
                {
                    worksheet.Cells[Row, 290 + new_row] = ListDocmuentos.ingreso_traspaso[i].Fecha;
                    worksheet.Cells[Row, 291 + new_row] = ListDocmuentos.ingreso_traspaso[i].Serie;
                    worksheet.Cells[Row, 292 + new_row] = ListDocmuentos.ingreso_traspaso[i].Folio;
                    worksheet.Cells[Row, 293 + new_row] = ListDocmuentos.ingreso_traspaso[i].NombreAgente;
                    worksheet.Cells[Row, 294 + new_row] = ListDocmuentos.ingreso_traspaso[i].RazonSocial;
                    worksheet.Cells[Row, 295 + new_row] = ListDocmuentos.ingreso_traspaso[i].FechaVencimiento;
                    worksheet.Cells[Row, 296 + new_row] = ListDocmuentos.ingreso_traspaso[i].TextoExtra1;
                    worksheet.Cells[Row, 297 + new_row] = ListDocmuentos.ingreso_traspaso[i].RFC;
                    worksheet.Cells[Row, 298 + new_row] = ListDocmuentos.ingreso_traspaso[i].Total;
                    worksheet.Cells[Row, 299 + new_row] = ListDocmuentos.ingreso_traspaso[i].Pendiente;
                    worksheet.Cells[Row, 300 + new_row] = ListDocmuentos.ingreso_traspaso[i].TextoExtra2;
                    worksheet.Cells[Row, 301 + new_row] = ListDocmuentos.ingreso_traspaso[i].Referencia;
                    total += ListDocmuentos.ingreso_traspaso[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 298 + new_row] = "$ " + total;
                #endregion
                #endregion

                //#region Ingreso dev. garantia
                //Row = 4;
                //#region encabezados
                //worksheet.Cells[2, 309 + new_row] = "Ingreso Dev. garantía";

                ////encabezados ingtreso traspaso
                //worksheet.Cells[Row, 305 + new_row] = "Fecha";
                //worksheet.Cells[Row, 306 + new_row] = "Serie";
                //worksheet.Cells[Row, 307 + new_row] = "Folio";
                //worksheet.Cells[Row, 308 + new_row] = "Nombre del agente";
                //worksheet.Cells[Row, 309 + new_row] = "Razon social";
                //worksheet.Cells[Row, 310 + new_row] = "Fecha de vencimiento";
                //worksheet.Cells[Row, 311 + new_row] = "Fecha de depósito";
                //worksheet.Cells[Row, 312 + new_row] = "RFC";
                //worksheet.Cells[Row, 313 + new_row] = "TOTAL";
                //worksheet.Cells[Row, 314 + new_row] = "Pendiente";
                //worksheet.Cells[Row, 315 + new_row] = "texto extra 2";
                //worksheet.Cells[Row, 316 + new_row] = "Referencia";

                //Row++;
                //#endregion
                //#region contenido
                //total = 0;
                //for (int i = 0; i < ListDocmuentos.ingreso_dev_garantia.Count; i++)
                //{
                //    worksheet.Cells[Row, 305 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].Fecha;
                //    worksheet.Cells[Row, 306 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].Serie;
                //    worksheet.Cells[Row, 307 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].Folio;
                //    worksheet.Cells[Row, 308 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].NombreAgente;
                //    worksheet.Cells[Row, 309 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].RazonSocial;
                //    worksheet.Cells[Row, 310 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].FechaVencimiento;
                //    worksheet.Cells[Row, 311 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].TextoExtra1;
                //    worksheet.Cells[Row, 312 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].RFC;
                //    worksheet.Cells[Row, 313 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].Total;
                //    worksheet.Cells[Row, 314 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].Pendiente;
                //    worksheet.Cells[Row, 315 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].TextoExtra2;
                //    worksheet.Cells[Row, 316 + new_row] = ListDocmuentos.ingreso_dev_garantia[i].Referencia;
                //    total += ListDocmuentos.ingreso_dev_garantia[i].Total;
                //    Row++;
                //}
                //worksheet.Cells[2, 313 + new_row] = "$ " + total;
                //#endregion
                //#endregion

            }
            catch (Exception)
            {

            }
        }




        public void excel_importCRUFletes(List<Tipos_Datos_CRU.show_fletes> ListDocmuentos)
        { //importar datos en excel
            try
            {


                // creating Excel Application
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;



                // see the excel sheet behind the program
                app.Visible = true;
                // get the reference of first sheet. By default its name is Sheet1.
                // store its reference to worksheet
                worksheet = workbook.Sheets["Hoja1"];
                worksheet = workbook.ActiveSheet;
                #region formato
                #region fletes
                Microsoft.Office.Interop.Excel.Range formatRange;
                formatRange = worksheet.get_Range("A4", "s1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion
                


                #endregion
                // changing the name of active sheet



                worksheet.Name = "Admipaq";
                int Row = 4;
                //titulo
                #region facturas
                Row = 4; //inicia a escribir en la fila 4
                #region encabezados
               // worksheet.Cells[2, 2] = "Fletes";

                //encabezados fletes
                worksheet.Cells[Row, 1] = "Fecha";
                worksheet.Cells[Row, 2] = "Folio";
                worksheet.Cells[Row, 3] = "Concepto";
                worksheet.Cells[Row, 4] = "Nombre Concepto";
                worksheet.Cells[Row, 5] = "Unidades";
                worksheet.Cells[Row, 6] = "Precio";
                worksheet.Cells[Row, 7] = "Total";
                worksheet.Cells[Row, 8] = "Cancelado";
                
                Row++;
                #endregion

                #region contenido
                float total = 0;
                for (int i = 0; i < ListDocmuentos.Count; i++)
                {
                    worksheet.Cells[Row, 1] = ListDocmuentos[i].fecha;
                    worksheet.Cells[Row, 2] = ListDocmuentos[i].Folio;
                    worksheet.Cells[Row, 3] = ListDocmuentos[i].concepto;
                    worksheet.Cells[Row, 4] = ListDocmuentos[i].nombre_concepto;
                    worksheet.Cells[Row, 5] = ListDocmuentos[i].Unidades;
                    worksheet.Cells[Row, 6] = ListDocmuentos[i].Precio;
                    worksheet.Cells[Row, 7] = ListDocmuentos[i].Total;
                    worksheet.Cells[Row, 8] = ListDocmuentos[i].Cancelado;

                    
                        total += ListDocmuentos[i].Total;

                    
                    Row++;
                }
                worksheet.Cells[2, 6] = "$ " + total;

                #endregion



                #endregion


            }
            catch (Exception)
            {

            }
        }
        #endregion 



        #region IMPRT EXCEL OL
        public void excel_importOL(Tipos_Datos_CRU.ListDatosOL ListDocmuentos)
        { //importar datos en excel
            try
            {


                // creating Excel Application
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // see the excel sheet behind the program
                app.Visible = true;
                // get the reference of first sheet. By default its name is Sheet1.
                // store its reference to worksheet
                worksheet = workbook.Sheets["Hoja1"];
                worksheet = workbook.ActiveSheet;
                // changing the name of active sheet
                worksheet.Name = "Admipaq";



                #region configuracion
                #region facturs
                Microsoft.Office.Interop.Excel.Range formatRange;
                formatRange = worksheet.get_Range("A4", "s1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion
                //
                formatRange = worksheet.get_Range("v4", "ao1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("AS4", "BK1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("BN4", "CF1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("CL4", "DD1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("DG4", "DY1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("EA4", "ES1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("EW4", "FO1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("FS4", "GK1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("GO4", "HH1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("HK4", "ID1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("IF4", "IX1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("JA4", "JS1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion
                int Row = 4;
                //titulo
                #region facturas
                Row = 4; //inicia a escribir en la fila 4
                #region encabezados
                worksheet.Cells[2, 5] = "facturas acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 1] = "Fecha";
                worksheet.Cells[Row, 2] = "Referencia";
                worksheet.Cells[Row, 3] = "Folio";
                worksheet.Cells[Row, 4] = "Nombre del agente";
                worksheet.Cells[Row, 5] = "Razon social";
                worksheet.Cells[Row, 6] = "Fecha de vencimiento";
                worksheet.Cells[Row, 7] = "RFC";
                worksheet.Cells[Row, 8] = "Subtotal";
                worksheet.Cells[Row, 9] = "IVA";
                worksheet.Cells[Row, 10] = "TOTAL";
                worksheet.Cells[Row, 11] = "Pendiente";
                worksheet.Cells[Row, 12] = "Texto Extra 3";
                worksheet.Cells[Row, 13] = "Afectado";
                worksheet.Cells[Row, 14] = "Impreso";
                worksheet.Cells[Row, 15] = "Cancelado";
                worksheet.Cells[Row, 16] = "Total de unidades";
                worksheet.Cells[Row, 17] = "Clasificacion cliente2";
                worksheet.Cells[Row, 18] = "Texto extra1";
                worksheet.Cells[Row, 19] = "Nombre del concepto";

                worksheet.Cells[2, 27] = "facturas público";

                //encabezados facturas
                worksheet.Cells[Row, 22] = "Fecha";
                worksheet.Cells[Row, 24] = "Referencia";
                worksheet.Cells[Row, 25] = "Folio";
                worksheet.Cells[Row, 26] = "Nombre del agente";
                worksheet.Cells[Row, 27] = "Razon social";
                worksheet.Cells[Row, 28] = "Fecha de vencimiento";
                worksheet.Cells[Row, 29] = "RFC";
                worksheet.Cells[Row, 30] = "Subtotal";
                worksheet.Cells[Row, 31] = "IVA";
                worksheet.Cells[Row, 32] = "TOTAL";
                worksheet.Cells[Row, 33] = "Pendiente";
                worksheet.Cells[Row, 34] = "Texto Extra 3";
                worksheet.Cells[Row, 35] = "Afectado";
                worksheet.Cells[Row, 36] = "Impreso";
                worksheet.Cells[Row, 37] = "Cancelado";
                worksheet.Cells[Row, 38] = "Total de unidades";
                worksheet.Cells[Row, 39] = "Clasificacion cliente2";
                worksheet.Cells[Row, 40] = "Texto extra1";
                worksheet.Cells[Row, 41] = "Nombre del concepto";

                worksheet.Cells[2, 49] = "facturas cliente por plaza";

                //encabezados facturas
                worksheet.Cells[Row, 45] = "Fecha";
                worksheet.Cells[Row, 46] = "Referencia";
                worksheet.Cells[Row, 47] = "Folio";
                worksheet.Cells[Row, 48] = "Nombre del agente";
                worksheet.Cells[Row, 49] = "Razon social";
                worksheet.Cells[Row, 50] = "Fecha de vencimiento";
                worksheet.Cells[Row, 51] = "RFC";
                worksheet.Cells[Row, 52] = "Subtotal";
                worksheet.Cells[Row, 53] = "IVA";
                worksheet.Cells[Row, 54] = "TOTAL";
                worksheet.Cells[Row, 55] = "Pendiente";
                worksheet.Cells[Row, 56] = "Texto Extra 3";
                worksheet.Cells[Row, 57] = "Afectado";
                worksheet.Cells[Row, 58] = "Impreso";
                worksheet.Cells[Row, 59] = "Cancelado";
                worksheet.Cells[Row, 60] = "Total de unidades";
                worksheet.Cells[Row, 61] = "Clasificacion cliente2";
                worksheet.Cells[Row, 62] = "Texto extra1";
                worksheet.Cells[Row, 63] = "Nombre del concepto";


                worksheet.Cells[2, 70] = "facturas por plazas";

                //encabezados facturas
                worksheet.Cells[Row, 66] = "Fecha";
                worksheet.Cells[Row, 67] = "Referencia";
                worksheet.Cells[Row, 68] = "Folio";
                worksheet.Cells[Row, 69] = "Nombre del agente";
                worksheet.Cells[Row, 70] = "Razon social";
                worksheet.Cells[Row, 71] = "Fecha de vencimiento";
                worksheet.Cells[Row, 72] = "RFC";
                worksheet.Cells[Row, 73] = "Subtotal";
                worksheet.Cells[Row, 74] = "IVA";
                worksheet.Cells[Row, 75] = "TOTAL";
                worksheet.Cells[Row, 76] = "Pendiente";
                worksheet.Cells[Row, 77] = "Texto Extra 3";
                worksheet.Cells[Row, 78] = "Afectado";
                worksheet.Cells[Row, 79] = "Impreso";
                worksheet.Cells[Row, 80] = "Cancelado";
                worksheet.Cells[Row, 81] = "Total de unidades";
                worksheet.Cells[Row, 82] = "Clasificacion cliente2";
                worksheet.Cells[Row, 83] = "Texto extra1";
                worksheet.Cells[Row, 84] = "Nombre del concepto";
                
                Row++;
                #endregion

                #region contenido
                float total = 0;
                Row = 5;
                float pendiente = 0;
                for (int i = 0; i < ListDocmuentos.facturas.Count; i++)
                {
                    worksheet.Cells[Row, 1] = ListDocmuentos.facturas[i].Fecha;
                    worksheet.Cells[Row, 2] = ListDocmuentos.facturas[i].Referencia;
                    worksheet.Cells[Row, 3] = ListDocmuentos.facturas[i].Folio;
                    worksheet.Cells[Row, 4] = ListDocmuentos.facturas[i].NombreAgente;
                    worksheet.Cells[Row, 5] = ListDocmuentos.facturas[i].RazonSocial;
                    worksheet.Cells[Row, 6] = ListDocmuentos.facturas[i].FechaVencimiento;
                    worksheet.Cells[Row, 7] = ListDocmuentos.facturas[i].RFC;
                    worksheet.Cells[Row, 8] = ListDocmuentos.facturas[i].Subtotal;
                    worksheet.Cells[Row, 9] = ListDocmuentos.facturas[i].IVA;
                    worksheet.Cells[Row, 10] = ListDocmuentos.facturas[i].Total;
                    worksheet.Cells[Row, 11] = ListDocmuentos.facturas[i].Pendiente;
                    worksheet.Cells[Row, 12] = ListDocmuentos.facturas[i].TextoExtra3;
                    worksheet.Cells[Row, 13] = ListDocmuentos.facturas[i].Afectado;
                    worksheet.Cells[Row, 14] = ListDocmuentos.facturas[i].Impreso;
                    worksheet.Cells[Row, 15] = ListDocmuentos.facturas[i].Cancelado;
                    worksheet.Cells[Row, 16] = ListDocmuentos.facturas[i].TotalUnidades;
                    worksheet.Cells[Row, 17] = ListDocmuentos.facturas[i].Clasificacion2;
                    worksheet.Cells[Row, 18] = ListDocmuentos.facturas[i].TextoExtra1;
                    worksheet.Cells[Row, 19] = ListDocmuentos.facturas[i].NombreConcepto;
                    if (ListDocmuentos.facturas[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas[i].Total;
                        pendiente += ListDocmuentos.facturas[i].Pendiente;
                    }
                    Row++;
                }
                worksheet.Cells[2, 10] = "$ " + total;
                worksheet.Cells[2, 11] = "$ " + pendiente;

                total = 0;
                pendiente = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.facturas_publico.Count; i++)
                {
                    worksheet.Cells[Row, 23] = ListDocmuentos.facturas_publico[i].Fecha;
                    worksheet.Cells[Row, 24] = ListDocmuentos.facturas_publico[i].Referencia;
                    worksheet.Cells[Row, 25] = ListDocmuentos.facturas_publico[i].Folio;
                    worksheet.Cells[Row, 26] = ListDocmuentos.facturas_publico[i].NombreAgente;
                    worksheet.Cells[Row, 27] = ListDocmuentos.facturas_publico[i].RazonSocial;
                    worksheet.Cells[Row, 28] = ListDocmuentos.facturas_publico[i].FechaVencimiento;
                    worksheet.Cells[Row, 29] = ListDocmuentos.facturas_publico[i].RFC;
                    worksheet.Cells[Row, 30] = ListDocmuentos.facturas_publico[i].Subtotal;
                    worksheet.Cells[Row, 31] = ListDocmuentos.facturas_publico[i].IVA;
                    worksheet.Cells[Row, 32] = ListDocmuentos.facturas_publico[i].Total;
                    worksheet.Cells[Row, 33] = ListDocmuentos.facturas_publico[i].Pendiente;
                    worksheet.Cells[Row, 34] = ListDocmuentos.facturas_publico[i].TextoExtra3;
                    worksheet.Cells[Row, 35] = ListDocmuentos.facturas_publico[i].Afectado;
                    worksheet.Cells[Row, 36] = ListDocmuentos.facturas_publico[i].Impreso;
                    worksheet.Cells[Row, 37] = ListDocmuentos.facturas_publico[i].Cancelado;
                    worksheet.Cells[Row, 38] = ListDocmuentos.facturas_publico[i].TotalUnidades;
                    worksheet.Cells[Row, 39] = ListDocmuentos.facturas_publico[i].Clasificacion2;
                    worksheet.Cells[Row, 40] = ListDocmuentos.facturas_publico[i].TextoExtra1;
                    worksheet.Cells[Row, 41] = ListDocmuentos.facturas_publico[i].NombreConcepto;
                    if (ListDocmuentos.facturas_publico[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas_publico[i].Total;
                        pendiente += ListDocmuentos.facturas_publico[i].Pendiente;
                    }
                    Row++;
                }
                worksheet.Cells[2, 32] = "$ " + total;
                worksheet.Cells[2, 33] = "$ " + pendiente;

                total = 0;
                pendiente = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.facturas_clientes_plazas.Count; i++)
                {
                    worksheet.Cells[Row, 45] = ListDocmuentos.facturas_clientes_plazas[i].Fecha;
                    worksheet.Cells[Row, 46] = ListDocmuentos.facturas_clientes_plazas[i].Referencia;
                    worksheet.Cells[Row, 47] = ListDocmuentos.facturas_clientes_plazas[i].Folio;
                    worksheet.Cells[Row, 48] = ListDocmuentos.facturas_clientes_plazas[i].NombreAgente;
                    worksheet.Cells[Row, 49] = ListDocmuentos.facturas_clientes_plazas[i].RazonSocial;
                    worksheet.Cells[Row, 50] = ListDocmuentos.facturas_clientes_plazas[i].FechaVencimiento;
                    worksheet.Cells[Row, 51] = ListDocmuentos.facturas_clientes_plazas[i].RFC;
                    worksheet.Cells[Row, 52] = ListDocmuentos.facturas_clientes_plazas[i].Subtotal;
                    worksheet.Cells[Row, 53] = ListDocmuentos.facturas_clientes_plazas[i].IVA;
                    worksheet.Cells[Row, 54] = ListDocmuentos.facturas_clientes_plazas[i].Total;
                    worksheet.Cells[Row, 55] = ListDocmuentos.facturas_clientes_plazas[i].Pendiente;
                    worksheet.Cells[Row, 56] = ListDocmuentos.facturas_clientes_plazas[i].TextoExtra3;
                    worksheet.Cells[Row, 57] = ListDocmuentos.facturas_clientes_plazas[i].Afectado;
                    worksheet.Cells[Row, 58] = ListDocmuentos.facturas_clientes_plazas[i].Impreso;
                    worksheet.Cells[Row, 59] = ListDocmuentos.facturas_clientes_plazas[i].Cancelado;
                    worksheet.Cells[Row, 60] = ListDocmuentos.facturas_clientes_plazas[i].TotalUnidades;
                    worksheet.Cells[Row, 61] = ListDocmuentos.facturas_clientes_plazas[i].Clasificacion2;
                    worksheet.Cells[Row, 62] = ListDocmuentos.facturas_clientes_plazas[i].TextoExtra1;
                    worksheet.Cells[Row, 63] = ListDocmuentos.facturas_clientes_plazas[i].NombreConcepto;
                    if (ListDocmuentos.facturas_clientes_plazas[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas_clientes_plazas[i].Total;
                        pendiente += ListDocmuentos.facturas_clientes_plazas[i].Pendiente;
                    }
                    Row++;
                }
                worksheet.Cells[2, 54] = "$ " + total;
                worksheet.Cells[2, 55] = "$ " + pendiente;

                total = 0;
                pendiente = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.facturas_por_plazas.Count; i++)
                {
                    worksheet.Cells[Row, 66] = ListDocmuentos.facturas_por_plazas[i].Fecha;
                    worksheet.Cells[Row, 67] = ListDocmuentos.facturas_por_plazas[i].Referencia;
                    worksheet.Cells[Row, 68] = ListDocmuentos.facturas_por_plazas[i].Folio;
                    worksheet.Cells[Row, 69] = ListDocmuentos.facturas_por_plazas[i].NombreAgente;
                    worksheet.Cells[Row, 70] = ListDocmuentos.facturas_por_plazas[i].RazonSocial;
                    worksheet.Cells[Row, 71] = ListDocmuentos.facturas_por_plazas[i].FechaVencimiento;
                    worksheet.Cells[Row, 72] = ListDocmuentos.facturas_por_plazas[i].RFC;
                    worksheet.Cells[Row, 73] = ListDocmuentos.facturas_por_plazas[i].Subtotal;
                    worksheet.Cells[Row, 74] = ListDocmuentos.facturas_por_plazas[i].IVA;
                    worksheet.Cells[Row, 75] = ListDocmuentos.facturas_por_plazas[i].Total;
                    worksheet.Cells[Row, 76] = ListDocmuentos.facturas_por_plazas[i].Pendiente;
                    worksheet.Cells[Row, 77] = ListDocmuentos.facturas_por_plazas[i].TextoExtra3;
                    worksheet.Cells[Row, 78] = ListDocmuentos.facturas_por_plazas[i].Afectado;
                    worksheet.Cells[Row, 79] = ListDocmuentos.facturas_por_plazas[i].Impreso;
                    worksheet.Cells[Row, 80] = ListDocmuentos.facturas_por_plazas[i].Cancelado;
                    worksheet.Cells[Row, 81] = ListDocmuentos.facturas_por_plazas[i].TotalUnidades;
                    worksheet.Cells[Row, 82] = ListDocmuentos.facturas_por_plazas[i].Clasificacion2;
                    worksheet.Cells[Row, 83] = ListDocmuentos.facturas_por_plazas[i].TextoExtra1;
                    worksheet.Cells[Row, 84] = ListDocmuentos.facturas_por_plazas[i].NombreConcepto;
                    if (ListDocmuentos.facturas_por_plazas[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas_por_plazas[i].Total;
                        pendiente += ListDocmuentos.facturas_por_plazas[i].Pendiente;
                    }
                    Row++;
                }
                worksheet.Cells[2, 75] = "$ " + total;
                worksheet.Cells[2, 76] = "$ " + pendiente;
                
                #endregion



                #endregion


                #region ABONOS
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 94] = "acumulado de abono";

                //encabezados facturas
                worksheet.Cells[Row, 90] = "Fecha";
                worksheet.Cells[Row, 91] = "Referencia";
                worksheet.Cells[Row, 92] = "Folio";
                worksheet.Cells[Row, 93] = "Nombre del agente";
                worksheet.Cells[Row, 94] = "Razon social";
                worksheet.Cells[Row, 95] = "Fecha de vencimiento";
                worksheet.Cells[Row, 96] = "RFC";
                worksheet.Cells[Row, 97] = "Subtotal";
                worksheet.Cells[Row, 98] = "IVA";
                worksheet.Cells[Row, 99] = "TOTAL";
                worksheet.Cells[Row, 100] = "Pendiente";
                worksheet.Cells[Row, 101] = "Texto Extra 3";
                worksheet.Cells[Row, 102] = "Afectado";
                worksheet.Cells[Row, 103] = "Impreso";
                worksheet.Cells[Row, 104] = "Cancelado";
                worksheet.Cells[Row, 105] = "Total de unidades";
                worksheet.Cells[Row, 106] = "Clasificacion cliente2";
                worksheet.Cells[Row, 107] = "Texto extra1";
                worksheet.Cells[Row, 108] = "Nombre del concepto";


                worksheet.Cells[2, 115] = "abonos públicos";

                //encabezados facturas
                worksheet.Cells[Row, 111] = "Fecha";
                worksheet.Cells[Row, 112] = "Referencia";
                worksheet.Cells[Row, 113] = "Folio";
                worksheet.Cells[Row, 114] = "Nombre del agente";
                worksheet.Cells[Row, 115] = "Razon social";
                worksheet.Cells[Row, 116] = "Fecha de vencimiento";
                worksheet.Cells[Row, 117] = "RFC";
                worksheet.Cells[Row, 118] = "Subtotal";
                worksheet.Cells[Row, 119] = "IVA";
                worksheet.Cells[Row, 120] = "TOTAL";
                worksheet.Cells[Row, 121] = "Pendiente";
                worksheet.Cells[Row, 122] = "Texto Extra 3";
                worksheet.Cells[Row, 123] = "Afectado";
                worksheet.Cells[Row, 124] = "Impreso";
                worksheet.Cells[Row, 125] = "Cancelado";
                worksheet.Cells[Row, 126] = "Total de unidades";
                worksheet.Cells[Row, 127] = "Clasificacion cliente2";
                worksheet.Cells[Row, 128] = "Texto extra1";
                worksheet.Cells[Row, 129] = "Nombre del concepto";

                worksheet.Cells[2, 135] = "abonos por plazas";

                //encabezados facturas
                worksheet.Cells[Row, 131] = "Fecha";
                worksheet.Cells[Row, 132] = "Referencia";
                worksheet.Cells[Row, 133] = "Folio";
                worksheet.Cells[Row, 134] = "Nombre del agente";
                worksheet.Cells[Row, 135] = "Razon social";
                worksheet.Cells[Row, 136] = "Fecha de vencimiento";
                worksheet.Cells[Row, 137] = "RFC";
                worksheet.Cells[Row, 138] = "Subtotal";
                worksheet.Cells[Row, 139] = "IVA";
                worksheet.Cells[Row, 140] = "TOTAL";
                worksheet.Cells[Row, 141] = "Pendiente";
                worksheet.Cells[Row, 142] = "Texto Extra 3";
                worksheet.Cells[Row, 143] = "Afectado";
                worksheet.Cells[Row, 144] = "Impreso";
                worksheet.Cells[Row, 145] = "Cancelado";
                worksheet.Cells[Row, 146] = "Total de unidades";
                worksheet.Cells[Row, 147] = "Clasificacion cliente2";
                worksheet.Cells[Row, 148] = "Texto extra1";
                worksheet.Cells[Row, 149] = "Nombre del concepto";


                worksheet.Cells[2, 157] = "acumulado de compras";

                //encabezados facturas
                worksheet.Cells[Row, 153] = "Fecha";
                worksheet.Cells[Row, 154] = "Referencia";
                worksheet.Cells[Row, 155] = "Folio";
                worksheet.Cells[Row, 156] = "Nombre del agente";
                worksheet.Cells[Row, 157] = "Razon social";
                worksheet.Cells[Row, 158] = "Fecha de vencimiento";
                worksheet.Cells[Row, 159] = "RFC";
                worksheet.Cells[Row, 160] = "Subtotal";
                worksheet.Cells[Row, 161] = "IVA";
                worksheet.Cells[Row, 162] = "TOTAL";
                worksheet.Cells[Row, 163] = "Pendiente";
                worksheet.Cells[Row, 164] = "Texto Extra 3";
                worksheet.Cells[Row, 165] = "Afectado";
                worksheet.Cells[Row, 166] = "Impreso";
                worksheet.Cells[Row, 167] = "Cancelado";
                worksheet.Cells[Row, 168] = "Total de unidades";
                worksheet.Cells[Row, 169] = "Clasificacion cliente2";
                worksheet.Cells[Row, 170] = "Texto extra1";
                worksheet.Cells[Row, 171] = "Nombre del concepto";


                worksheet.Cells[2, 179] = "compras a CRU";

                //encabezados facturas
                worksheet.Cells[Row, 175] = "Fecha";
                worksheet.Cells[Row, 176] = "Referencia";
                worksheet.Cells[Row, 177] = "Folio";
                worksheet.Cells[Row, 178] = "Nombre del agente";
                worksheet.Cells[Row, 179] = "Razon social";
                worksheet.Cells[Row, 180] = "Fecha de vencimiento";
                worksheet.Cells[Row, 181] = "RFC";
                worksheet.Cells[Row, 182] = "Subtotal";
                worksheet.Cells[Row, 183] = "IVA";
                worksheet.Cells[Row, 184] = "TOTAL";
                worksheet.Cells[Row, 185] = "Pendiente";
                worksheet.Cells[Row, 186] = "Texto Extra 3";
                worksheet.Cells[Row, 187] = "Afectado";
                worksheet.Cells[Row, 188] = "Impreso";
                worksheet.Cells[Row, 189] = "Cancelado";
                worksheet.Cells[Row, 190] = "Total de unidades";
                worksheet.Cells[Row, 191] = "Clasificacion cliente2";
                worksheet.Cells[Row, 192] = "Texto extra1";
                worksheet.Cells[Row, 193] = "Nombre del concepto";

                worksheet.Cells[2, 201] = "compras a Manuel";

                //encabezados facturas
                worksheet.Cells[Row, 197] = "Fecha";
                worksheet.Cells[Row, 198] = "Referencia";
                worksheet.Cells[Row, 199] = "Folio";
                worksheet.Cells[Row, 200] = "Nombre del agente";
                worksheet.Cells[Row, 201] = "Razon social";
                worksheet.Cells[Row, 202] = "Fecha de vencimiento";
                worksheet.Cells[Row, 203] = "RFC";
                worksheet.Cells[Row, 204] = "Subtotal";
                worksheet.Cells[Row, 205] = "IVA";
                worksheet.Cells[Row, 206] = "TOTAL";
                worksheet.Cells[Row, 207] = "Pendiente";
                worksheet.Cells[Row, 208] = "Texto Extra 3";
                worksheet.Cells[Row, 209] = "Afectado";
                worksheet.Cells[Row, 210] = "Impreso";
                worksheet.Cells[Row, 211] = "Cancelado";
                worksheet.Cells[Row, 212] = "Total de unidades";
                worksheet.Cells[Row, 213] = "Clasificacion cliente2";
                worksheet.Cells[Row, 214] = "Texto extra1";
                worksheet.Cells[Row, 215] = "Nombre del concepto";


                worksheet.Cells[2, 223] = "pagos al proveedor";

                //encabezados facturas
                worksheet.Cells[Row, 219] = "Fecha";
                worksheet.Cells[Row, 220] = "Referencia";
                worksheet.Cells[Row, 221] = "Folio";
                worksheet.Cells[Row, 222] = "Nombre del agente";
                worksheet.Cells[Row, 223] = "Razon social";
                worksheet.Cells[Row, 224] = "Fecha de vencimiento";
                worksheet.Cells[Row, 225] = "RFC";
                worksheet.Cells[Row, 226] = "Subtotal";
                worksheet.Cells[Row, 227] = "IVA";
                worksheet.Cells[Row, 228] = "TOTAL";
                worksheet.Cells[Row, 229] = "Pendiente";
                worksheet.Cells[Row, 230] = "Texto Extra 3";
                worksheet.Cells[Row, 231] = "Afectado";
                worksheet.Cells[Row, 232] = "Impreso";
                worksheet.Cells[Row, 233] = "Cancelado";
                worksheet.Cells[Row, 234] = "Total de unidades";
                worksheet.Cells[Row, 235] = "Clasificacion cliente2";
                worksheet.Cells[Row, 236] = "Texto extra1";
                worksheet.Cells[Row, 237] = "Nombre del concepto";


                worksheet.Cells[2, 244] = "pagos al proveedor a CRU";

                //encabezados facturas
                worksheet.Cells[Row, 240] = "Fecha";
                worksheet.Cells[Row, 241] = "Referencia";
                worksheet.Cells[Row, 242] = "Folio";
                worksheet.Cells[Row, 243] = "Nombre del agente";
                worksheet.Cells[Row, 244] = "Razon social";
                worksheet.Cells[Row, 245] = "Fecha de vencimiento";
                worksheet.Cells[Row, 246] = "RFC";
                worksheet.Cells[Row, 247] = "Subtotal";
                worksheet.Cells[Row, 248] = "IVA";
                worksheet.Cells[Row, 249] = "TOTAL";
                worksheet.Cells[Row, 250] = "Pendiente";
                worksheet.Cells[Row, 251] = "Texto Extra 3";
                worksheet.Cells[Row, 252] = "Afectado";
                worksheet.Cells[Row, 253] = "Impreso";
                worksheet.Cells[Row, 254] = "Cancelado";
                worksheet.Cells[Row, 255] = "Total de unidades";
                worksheet.Cells[Row, 256] = "Clasificacion cliente2";
                worksheet.Cells[Row, 257] = "Texto extra1";
                worksheet.Cells[Row, 258] = "Nombre del concepto";


                worksheet.Cells[2, 265] = "pagos al proveedor a Manuel";

                //encabezados facturas
                worksheet.Cells[Row, 261] = "Fecha";
                worksheet.Cells[Row, 262] = "Referencia";
                worksheet.Cells[Row, 263] = "Folio";
                worksheet.Cells[Row, 264] = "Nombre del agente";
                worksheet.Cells[Row, 265] = "Razon social";
                worksheet.Cells[Row, 266] = "Fecha de vencimiento";
                worksheet.Cells[Row, 267] = "RFC";
                worksheet.Cells[Row, 268] = "Subtotal";
                worksheet.Cells[Row, 269] = "IVA";
                worksheet.Cells[Row, 270] = "TOTAL";
                worksheet.Cells[Row, 271] = "Pendiente";
                worksheet.Cells[Row, 272] = "Texto Extra 3";
                worksheet.Cells[Row, 273] = "Afectado";
                worksheet.Cells[Row, 274] = "Impreso";
                worksheet.Cells[Row, 275] = "Cancelado";
                worksheet.Cells[Row, 276] = "Total de unidades";
                worksheet.Cells[Row, 277] = "Clasificacion cliente2";
                worksheet.Cells[Row, 278] = "Texto extra1";
                worksheet.Cells[Row, 279] = "Nombre del concepto";
                Row++;
                #endregion
                #region contenido
                total = 0;
                Row = 5;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.abonos.Count; i++)
                {
                    worksheet.Cells[Row, 90] = ListDocmuentos.abonos[i].Fecha;
                    worksheet.Cells[Row, 91] = ListDocmuentos.abonos[i].Referencia;
                    worksheet.Cells[Row, 92] = ListDocmuentos.abonos[i].Folio;
                    worksheet.Cells[Row, 93] = ListDocmuentos.abonos[i].NombreAgente;
                    worksheet.Cells[Row, 94] = ListDocmuentos.abonos[i].RazonSocial;
                    worksheet.Cells[Row, 95] = ListDocmuentos.abonos[i].FechaVencimiento;
                    worksheet.Cells[Row, 96] = ListDocmuentos.abonos[i].RFC;
                    worksheet.Cells[Row, 97] = ListDocmuentos.abonos[i].Subtotal;
                    worksheet.Cells[Row, 98] = ListDocmuentos.abonos[i].IVA;
                    worksheet.Cells[Row, 99] = ListDocmuentos.abonos[i].Total;
                    worksheet.Cells[Row, 100] = ListDocmuentos.abonos[i].Pendiente;
                    worksheet.Cells[Row, 101] = ListDocmuentos.abonos[i].TextoExtra3;
                    worksheet.Cells[Row, 102] = ListDocmuentos.abonos[i].Afectado;
                    worksheet.Cells[Row, 103] = ListDocmuentos.abonos[i].Impreso;
                    worksheet.Cells[Row, 104] = ListDocmuentos.abonos[i].Cancelado;
                    worksheet.Cells[Row, 105] = ListDocmuentos.abonos[i].TotalUnidades;
                    worksheet.Cells[Row, 106] = ListDocmuentos.abonos[i].Clasificacion2;
                    worksheet.Cells[Row, 107] = ListDocmuentos.abonos[i].TextoExtra1;
                    worksheet.Cells[Row, 108] = ListDocmuentos.abonos[i].NombreConcepto;
                    total += ListDocmuentos.abonos[i].Total;
                    pendiente += ListDocmuentos.abonos[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 99] = "$ " + total;
                worksheet.Cells[2, 100] = "$ " + pendiente;

                Row = 5;
                total = 0;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.abonos_publico.Count; i++)
                {
                    worksheet.Cells[Row, 111] = ListDocmuentos.abonos_publico[i].Fecha;
                    worksheet.Cells[Row, 112] = ListDocmuentos.abonos_publico[i].Referencia;
                    worksheet.Cells[Row, 113] = ListDocmuentos.abonos_publico[i].Folio;
                    worksheet.Cells[Row, 114] = ListDocmuentos.abonos_publico[i].NombreAgente;
                    worksheet.Cells[Row, 115] = ListDocmuentos.abonos_publico[i].RazonSocial;
                    worksheet.Cells[Row, 116] = ListDocmuentos.abonos_publico[i].FechaVencimiento;
                    worksheet.Cells[Row, 117] = ListDocmuentos.abonos_publico[i].RFC;
                    worksheet.Cells[Row, 118] = ListDocmuentos.abonos_publico[i].Subtotal;
                    worksheet.Cells[Row, 119] = ListDocmuentos.abonos_publico[i].IVA;
                    worksheet.Cells[Row, 120] = ListDocmuentos.abonos_publico[i].Total;
                    worksheet.Cells[Row, 121] = ListDocmuentos.abonos_publico[i].Pendiente;
                    worksheet.Cells[Row, 122] = ListDocmuentos.abonos_publico[i].TextoExtra3;
                    worksheet.Cells[Row, 123] = ListDocmuentos.abonos_publico[i].Afectado;
                    worksheet.Cells[Row, 124] = ListDocmuentos.abonos_publico[i].Impreso;
                    worksheet.Cells[Row, 125] = ListDocmuentos.abonos_publico[i].Cancelado;
                    worksheet.Cells[Row, 126] = ListDocmuentos.abonos_publico[i].TotalUnidades;
                    worksheet.Cells[Row, 127] = ListDocmuentos.abonos_publico[i].Clasificacion2;
                    worksheet.Cells[Row, 128] = ListDocmuentos.abonos_publico[i].TextoExtra1;
                    worksheet.Cells[Row, 129] = ListDocmuentos.abonos_publico[i].NombreConcepto;
                    total += ListDocmuentos.abonos_publico[i].Total;
                    pendiente += ListDocmuentos.abonos_publico[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 120] = "$ " + total;
                worksheet.Cells[2, 121] = "$ " + pendiente;

                Row = 5;
                total = 0;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.abonos_plazas.Count; i++)
                {
                    worksheet.Cells[Row, 131] = ListDocmuentos.abonos_plazas[i].Fecha;
                    worksheet.Cells[Row, 132] = ListDocmuentos.abonos_plazas[i].Referencia;
                    worksheet.Cells[Row, 133] = ListDocmuentos.abonos_plazas[i].Folio;
                    worksheet.Cells[Row, 134] = ListDocmuentos.abonos_plazas[i].NombreAgente;
                    worksheet.Cells[Row, 135] = ListDocmuentos.abonos_plazas[i].RazonSocial;
                    worksheet.Cells[Row, 136] = ListDocmuentos.abonos_plazas[i].FechaVencimiento;
                    worksheet.Cells[Row, 137] = ListDocmuentos.abonos_plazas[i].RFC;
                    worksheet.Cells[Row, 138] = ListDocmuentos.abonos_plazas[i].Subtotal;
                    worksheet.Cells[Row, 139] = ListDocmuentos.abonos_plazas[i].IVA;
                    worksheet.Cells[Row, 140] = ListDocmuentos.abonos_plazas[i].Total;
                    worksheet.Cells[Row, 141] = ListDocmuentos.abonos_plazas[i].Pendiente;
                    worksheet.Cells[Row, 142] = ListDocmuentos.abonos_plazas[i].TextoExtra3;
                    worksheet.Cells[Row, 143] = ListDocmuentos.abonos_plazas[i].Afectado;
                    worksheet.Cells[Row, 144] = ListDocmuentos.abonos_plazas[i].Impreso;
                    worksheet.Cells[Row, 145] = ListDocmuentos.abonos_plazas[i].Cancelado;
                    worksheet.Cells[Row, 146] = ListDocmuentos.abonos_plazas[i].TotalUnidades;
                    worksheet.Cells[Row, 147] = ListDocmuentos.abonos_plazas[i].Clasificacion2;
                    worksheet.Cells[Row, 148] = ListDocmuentos.abonos_plazas[i].TextoExtra1;
                    worksheet.Cells[Row, 149] = ListDocmuentos.abonos_plazas[i].NombreConcepto;
                    total += ListDocmuentos.abonos_plazas[i].Total;
                    pendiente += ListDocmuentos.abonos_plazas[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 140] = "$ " + total;
                worksheet.Cells[2, 141] = "$ " + pendiente;

                Row = 5;
                total = 0;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.compras.Count; i++)
                {
                    worksheet.Cells[Row, 153] = ListDocmuentos.compras[i].Fecha;
                    worksheet.Cells[Row, 154] = ListDocmuentos.compras[i].Referencia;
                    worksheet.Cells[Row, 155] = ListDocmuentos.compras[i].Folio;
                    worksheet.Cells[Row, 156] = ListDocmuentos.compras[i].NombreAgente;
                    worksheet.Cells[Row, 157] = ListDocmuentos.compras[i].RazonSocial;
                    worksheet.Cells[Row, 158] = ListDocmuentos.compras[i].FechaVencimiento;
                    worksheet.Cells[Row, 159] = ListDocmuentos.compras[i].RFC;
                    worksheet.Cells[Row, 160] = ListDocmuentos.compras[i].Subtotal;
                    worksheet.Cells[Row, 161] = ListDocmuentos.compras[i].IVA;
                    worksheet.Cells[Row, 162] = ListDocmuentos.compras[i].Total;
                    worksheet.Cells[Row, 163] = ListDocmuentos.compras[i].Pendiente;
                    worksheet.Cells[Row, 164] = ListDocmuentos.compras[i].TextoExtra3;
                    worksheet.Cells[Row, 165] = ListDocmuentos.compras[i].Afectado;
                    worksheet.Cells[Row, 166] = ListDocmuentos.compras[i].Impreso;
                    worksheet.Cells[Row, 167] = ListDocmuentos.compras[i].Cancelado;
                    worksheet.Cells[Row, 168] = ListDocmuentos.compras[i].TotalUnidades;
                    worksheet.Cells[Row, 169] = ListDocmuentos.compras[i].Clasificacion2;
                    worksheet.Cells[Row, 170] = ListDocmuentos.compras[i].TextoExtra1;
                    worksheet.Cells[Row, 171] = ListDocmuentos.compras[i].NombreConcepto;
                    total += ListDocmuentos.compras[i].Total;
                    pendiente += ListDocmuentos.compras[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 162] = "$ " + total;
                worksheet.Cells[2, 163] = "$ " + pendiente;
                Row = 5;
                total = 0;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.compras_cru.Count; i++)
                {
                    worksheet.Cells[Row, 175] = ListDocmuentos.compras_cru[i].Fecha;
                    worksheet.Cells[Row, 176] = ListDocmuentos.compras_cru[i].Referencia;
                    worksheet.Cells[Row, 177] = ListDocmuentos.compras_cru[i].Folio;
                    worksheet.Cells[Row, 178] = ListDocmuentos.compras_cru[i].NombreAgente;
                    worksheet.Cells[Row, 179] = ListDocmuentos.compras_cru[i].RazonSocial;
                    worksheet.Cells[Row, 180] = ListDocmuentos.compras_cru[i].FechaVencimiento;
                    worksheet.Cells[Row, 181] = ListDocmuentos.compras_cru[i].RFC;
                    worksheet.Cells[Row, 182] = ListDocmuentos.compras_cru[i].Subtotal;
                    worksheet.Cells[Row, 183] = ListDocmuentos.compras_cru[i].IVA;
                    worksheet.Cells[Row, 184] = ListDocmuentos.compras_cru[i].Total;
                    worksheet.Cells[Row, 185] = ListDocmuentos.compras_cru[i].Pendiente;
                    worksheet.Cells[Row, 186] = ListDocmuentos.compras_cru[i].TextoExtra3;
                    worksheet.Cells[Row, 187] = ListDocmuentos.compras_cru[i].Afectado;
                    worksheet.Cells[Row, 188] = ListDocmuentos.compras_cru[i].Impreso;
                    worksheet.Cells[Row, 189] = ListDocmuentos.compras_cru[i].Cancelado;
                    worksheet.Cells[Row, 190] = ListDocmuentos.compras_cru[i].TotalUnidades;
                    worksheet.Cells[Row, 191] = ListDocmuentos.compras_cru[i].Clasificacion2;
                    worksheet.Cells[Row, 192] = ListDocmuentos.compras_cru[i].TextoExtra1;
                    worksheet.Cells[Row, 193] = ListDocmuentos.compras_cru[i].NombreConcepto;
                    total += ListDocmuentos.compras_cru[i].Total;
                    pendiente += ListDocmuentos.compras_cru[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 184] = "$ " + total;
                worksheet.Cells[2, 185] = "$ " + pendiente;
                Row = 5;
                total = 0;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.compras_manuel.Count; i++)
                {
                    worksheet.Cells[Row, 197] = ListDocmuentos.compras_manuel[i].Fecha;
                    worksheet.Cells[Row, 198] = ListDocmuentos.compras_manuel[i].Referencia;
                    worksheet.Cells[Row, 199] = ListDocmuentos.compras_manuel[i].Folio;
                    worksheet.Cells[Row, 200] = ListDocmuentos.compras_manuel[i].NombreAgente;
                    worksheet.Cells[Row, 201] = ListDocmuentos.compras_manuel[i].RazonSocial;
                    worksheet.Cells[Row, 202] = ListDocmuentos.compras_manuel[i].FechaVencimiento;
                    worksheet.Cells[Row, 203] = ListDocmuentos.compras_manuel[i].RFC;
                    worksheet.Cells[Row, 204] = ListDocmuentos.compras_manuel[i].Subtotal;
                    worksheet.Cells[Row, 205] = ListDocmuentos.compras_manuel[i].IVA;
                    worksheet.Cells[Row, 206] = ListDocmuentos.compras_manuel[i].Total;
                    worksheet.Cells[Row, 207] = ListDocmuentos.compras_manuel[i].Pendiente;
                    worksheet.Cells[Row, 208] = ListDocmuentos.compras_manuel[i].TextoExtra3;
                    worksheet.Cells[Row, 209] = ListDocmuentos.compras_manuel[i].Afectado;
                    worksheet.Cells[Row, 210] = ListDocmuentos.compras_manuel[i].Impreso;
                    worksheet.Cells[Row, 211] = ListDocmuentos.compras_manuel[i].Cancelado;
                    worksheet.Cells[Row, 212] = ListDocmuentos.compras_manuel[i].TotalUnidades;
                    worksheet.Cells[Row, 213] = ListDocmuentos.compras_manuel[i].Clasificacion2;
                    worksheet.Cells[Row, 214] = ListDocmuentos.compras_manuel[i].TextoExtra1;
                    worksheet.Cells[Row, 215] = ListDocmuentos.compras_manuel[i].NombreConcepto;
                    total += ListDocmuentos.compras_manuel[i].Total;
                    pendiente += ListDocmuentos.compras_manuel[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 206] = "$ " + total;
                worksheet.Cells[2, 207] = "$ " + pendiente;
                Row = 5;
                total = 0;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.pagos_proveedor.Count; i++)
                {
                    worksheet.Cells[Row, 219] = ListDocmuentos.pagos_proveedor[i].Fecha;
                    worksheet.Cells[Row, 220] = ListDocmuentos.pagos_proveedor[i].Referencia;
                    worksheet.Cells[Row, 221] = ListDocmuentos.pagos_proveedor[i].Folio;
                    worksheet.Cells[Row, 222] = ListDocmuentos.pagos_proveedor[i].NombreAgente;
                    worksheet.Cells[Row, 223] = ListDocmuentos.pagos_proveedor[i].RazonSocial;
                    worksheet.Cells[Row, 224] = ListDocmuentos.pagos_proveedor[i].FechaVencimiento;
                    worksheet.Cells[Row, 225] = ListDocmuentos.pagos_proveedor[i].RFC;
                    worksheet.Cells[Row, 226] = ListDocmuentos.pagos_proveedor[i].Subtotal;
                    worksheet.Cells[Row, 227] = ListDocmuentos.pagos_proveedor[i].IVA;
                    worksheet.Cells[Row, 228] = ListDocmuentos.pagos_proveedor[i].Total;
                    worksheet.Cells[Row, 229] = ListDocmuentos.pagos_proveedor[i].Pendiente;
                    worksheet.Cells[Row, 230] = ListDocmuentos.pagos_proveedor[i].TextoExtra3;
                    worksheet.Cells[Row, 231] = ListDocmuentos.pagos_proveedor[i].Afectado;
                    worksheet.Cells[Row, 232] = ListDocmuentos.pagos_proveedor[i].Impreso;
                    worksheet.Cells[Row, 233] = ListDocmuentos.pagos_proveedor[i].Cancelado;
                    worksheet.Cells[Row, 234] = ListDocmuentos.pagos_proveedor[i].TotalUnidades;
                    worksheet.Cells[Row, 235] = ListDocmuentos.pagos_proveedor[i].Clasificacion2;
                    worksheet.Cells[Row, 236] = ListDocmuentos.pagos_proveedor[i].TextoExtra1;
                    worksheet.Cells[Row, 237] = ListDocmuentos.pagos_proveedor[i].NombreConcepto;
                    total += ListDocmuentos.pagos_proveedor[i].Total;
                    pendiente += ListDocmuentos.pagos_proveedor[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 228] = "$ " + total;
                worksheet.Cells[2, 229] = "$ " + pendiente;
                Row = 5;
                total = 0;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.pagos_proveedor_cru.Count; i++)
                {
                    worksheet.Cells[Row, 240] = ListDocmuentos.pagos_proveedor_cru[i].Fecha;
                    worksheet.Cells[Row, 241] = ListDocmuentos.pagos_proveedor_cru[i].Referencia;
                    worksheet.Cells[Row, 242] = ListDocmuentos.pagos_proveedor_cru[i].Folio;
                    worksheet.Cells[Row, 243] = ListDocmuentos.pagos_proveedor_cru[i].NombreAgente;
                    worksheet.Cells[Row, 244] = ListDocmuentos.pagos_proveedor_cru[i].RazonSocial;
                    worksheet.Cells[Row, 245] = ListDocmuentos.pagos_proveedor_cru[i].FechaVencimiento;
                    worksheet.Cells[Row, 246] = ListDocmuentos.pagos_proveedor_cru[i].RFC;
                    worksheet.Cells[Row, 247] = ListDocmuentos.pagos_proveedor_cru[i].Subtotal;
                    worksheet.Cells[Row, 248] = ListDocmuentos.pagos_proveedor_cru[i].IVA;
                    worksheet.Cells[Row, 249] = ListDocmuentos.pagos_proveedor_cru[i].Total;
                    worksheet.Cells[Row, 250] = ListDocmuentos.pagos_proveedor_cru[i].Pendiente;
                    worksheet.Cells[Row, 251] = ListDocmuentos.pagos_proveedor_cru[i].TextoExtra3;
                    worksheet.Cells[Row, 252] = ListDocmuentos.pagos_proveedor_cru[i].Afectado;
                    worksheet.Cells[Row, 253] = ListDocmuentos.pagos_proveedor_cru[i].Impreso;
                    worksheet.Cells[Row, 254] = ListDocmuentos.pagos_proveedor_cru[i].Cancelado;
                    worksheet.Cells[Row, 255] = ListDocmuentos.pagos_proveedor_cru[i].TotalUnidades;
                    worksheet.Cells[Row, 256] = ListDocmuentos.pagos_proveedor_cru[i].Clasificacion2;
                    worksheet.Cells[Row, 257] = ListDocmuentos.pagos_proveedor_cru[i].TextoExtra1;
                    worksheet.Cells[Row, 258] = ListDocmuentos.pagos_proveedor_cru[i].NombreConcepto;
                    total += ListDocmuentos.pagos_proveedor_cru[i].Total;
                    pendiente += ListDocmuentos.pagos_proveedor_cru[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 249] = "$ " + total;
                worksheet.Cells[2, 250] = "$ " + pendiente;
                Row = 5;
                total = 0;
                pendiente = 0;
                for (int i = 0; i < ListDocmuentos.pagos_proveedor_manuel.Count; i++)
                {
                    worksheet.Cells[Row, 261] = ListDocmuentos.pagos_proveedor_manuel[i].Fecha;
                    worksheet.Cells[Row, 262] = ListDocmuentos.pagos_proveedor_manuel[i].Referencia;
                    worksheet.Cells[Row, 263] = ListDocmuentos.pagos_proveedor_manuel[i].Folio;
                    worksheet.Cells[Row, 264] = ListDocmuentos.pagos_proveedor_manuel[i].NombreAgente;
                    worksheet.Cells[Row, 265] = ListDocmuentos.pagos_proveedor_manuel[i].RazonSocial;
                    worksheet.Cells[Row, 266] = ListDocmuentos.pagos_proveedor_manuel[i].FechaVencimiento;
                    worksheet.Cells[Row, 267] = ListDocmuentos.pagos_proveedor_manuel[i].RFC;
                    worksheet.Cells[Row, 268] = ListDocmuentos.pagos_proveedor_manuel[i].Subtotal;
                    worksheet.Cells[Row, 269] = ListDocmuentos.pagos_proveedor_manuel[i].IVA;
                    worksheet.Cells[Row, 270] = ListDocmuentos.pagos_proveedor_manuel[i].Total;
                    worksheet.Cells[Row, 271] = ListDocmuentos.pagos_proveedor_manuel[i].Pendiente;
                    worksheet.Cells[Row, 272] = ListDocmuentos.pagos_proveedor_manuel[i].TextoExtra3;
                    worksheet.Cells[Row, 273] = ListDocmuentos.pagos_proveedor_manuel[i].Afectado;
                    worksheet.Cells[Row, 274] = ListDocmuentos.pagos_proveedor_manuel[i].Impreso;
                    worksheet.Cells[Row, 275] = ListDocmuentos.pagos_proveedor_manuel[i].Cancelado;
                    worksheet.Cells[Row, 276] = ListDocmuentos.pagos_proveedor_manuel[i].TotalUnidades;
                    worksheet.Cells[Row, 277] = ListDocmuentos.pagos_proveedor_manuel[i].Clasificacion2;
                    worksheet.Cells[Row, 278] = ListDocmuentos.pagos_proveedor_manuel[i].TextoExtra1;
                    worksheet.Cells[Row, 279] = ListDocmuentos.pagos_proveedor_manuel[i].NombreConcepto;
                    total += ListDocmuentos.pagos_proveedor_manuel[i].Total;
                    pendiente += ListDocmuentos.pagos_proveedor_manuel[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 270] = "$ " + total;
                worksheet.Cells[2, 271] = "$ " + pendiente;
                #endregion
                #endregion


            }
            catch (Exception)
            {

            }
        }

        #endregion 
    
    
        #region IMPRT EXCEL ISEL
        public void excel_importISEL(Tipos_Datos_CRU.ListDatosISEL ListDocmuentos)
        { //importar datos en excel
            try
            {


                // creating Excel Application
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // see the excel sheet behind the program
                app.Visible = true;
                // get the reference of first sheet. By default its name is Sheet1.
                // store its reference to worksheet
                worksheet = workbook.Sheets["Hoja1"];
                worksheet = workbook.ActiveSheet;
                // changing the name of active sheet
                worksheet.Name = "Admipaq";
                #region configuracopn
                Microsoft.Office.Interop.Excel.Range formatRange;
                formatRange = worksheet.get_Range("A4", "s1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("V4", "AO1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("AS4", "BK1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("BM4", "CE1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna

                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                #endregion

                int Row = 4;
                //titulo
                #region facturas
                Row = 4; //inicia a escribir en la fila 4
                #region encabezados
                worksheet.Cells[2, 5] = "facturas acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 1] = "Fecha";
                worksheet.Cells[Row, 2] = "Referencia";
                worksheet.Cells[Row, 3] = "Folio";
                worksheet.Cells[Row, 4] = "Nombre del agente";
                worksheet.Cells[Row, 5] = "Razon social";
                worksheet.Cells[Row, 6] = "Fecha de vencimiento";
                worksheet.Cells[Row, 7] = "RFC";
                worksheet.Cells[Row, 8] = "Subtotal";
                worksheet.Cells[Row, 9] = "IVA";
                worksheet.Cells[Row, 10] = "TOTAL";
                worksheet.Cells[Row, 11] = "Pendiente";
                worksheet.Cells[Row, 12] = "Texto Extra 3";
                worksheet.Cells[Row, 13] = "Afectado";
                worksheet.Cells[Row, 14] = "Impreso";
                worksheet.Cells[Row, 15] = "Cancelado";
                worksheet.Cells[Row, 16] = "Total de unidades";
                worksheet.Cells[Row, 17] = "Clasificacion cliente2";
                worksheet.Cells[Row, 18] = "Texto extra1";
                worksheet.Cells[Row, 19] = "Nombre del concepto";

                worksheet.Cells[2, 27] = "facturas Dario";

                //encabezados facturas
                worksheet.Cells[Row, 22] = "Fecha";
                worksheet.Cells[Row, 24] = "Referencia";
                worksheet.Cells[Row, 25] = "Folio";
                worksheet.Cells[Row, 26] = "Nombre del agente";
                worksheet.Cells[Row, 27] = "Razon social";
                worksheet.Cells[Row, 28] = "Fecha de vencimiento";
                worksheet.Cells[Row, 29] = "RFC";
                worksheet.Cells[Row, 30] = "Subtotal";
                worksheet.Cells[Row, 31] = "IVA";
                worksheet.Cells[Row, 32] = "TOTAL";
                worksheet.Cells[Row, 33] = "Pendiente";
                worksheet.Cells[Row, 34] = "Texto Extra 3";
                worksheet.Cells[Row, 35] = "Afectado";
                worksheet.Cells[Row, 36] = "Impreso";
                worksheet.Cells[Row, 37] = "Cancelado";
                worksheet.Cells[Row, 38] = "Total de unidades";
                worksheet.Cells[Row, 39] = "Clasificacion cliente2";
                worksheet.Cells[Row, 40] = "Texto extra1";
                worksheet.Cells[Row, 41] = "Nombre del concepto";




                worksheet.Cells[2,49] = "abonos acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 45] = "Fecha";
                worksheet.Cells[Row, 46] = "Referencia";
                worksheet.Cells[Row, 47] = "Folio";
                worksheet.Cells[Row, 48] = "Nombre del agente";
                worksheet.Cells[Row, 49] = "Razon social";
                worksheet.Cells[Row, 50] = "Fecha de vencimiento";
                worksheet.Cells[Row, 51] = "RFC";
                worksheet.Cells[Row, 52] = "Subtotal";
                worksheet.Cells[Row, 53] = "IVA";
                worksheet.Cells[Row, 54] = "TOTAL";
                worksheet.Cells[Row, 55] = "Pendiente";
                worksheet.Cells[Row, 56] = "Texto Extra 3";
                worksheet.Cells[Row, 57] = "Afectado";
                worksheet.Cells[Row, 58] = "Impreso";
                worksheet.Cells[Row, 59] = "Cancelado";
                worksheet.Cells[Row, 60] = "Total de unidades";
                worksheet.Cells[Row, 61] = "Clasificacion cliente2";
                worksheet.Cells[Row, 62] = "Texto extra1";
                worksheet.Cells[Row, 63] = "Nombre del concepto";

                worksheet.Cells[2, 69] = "abonos Dario";

                //encabezados facturas
                worksheet.Cells[Row, 65] = "Fecha";
                worksheet.Cells[Row, 66] = "Referencia";
                worksheet.Cells[Row, 67] = "Folio";
                worksheet.Cells[Row, 68] = "Nombre del agente";
                worksheet.Cells[Row, 69] = "Razon social";
                worksheet.Cells[Row, 70] = "Fecha de vencimiento";
                worksheet.Cells[Row, 71] = "RFC";
                worksheet.Cells[Row, 72] = "Subtotal";
                worksheet.Cells[Row, 73] = "IVA";
                worksheet.Cells[Row, 74] = "TOTAL";
                worksheet.Cells[Row, 75] = "Pendiente";
                worksheet.Cells[Row, 76] = "Texto Extra 3";
                worksheet.Cells[Row, 77] = "Afectado";
                worksheet.Cells[Row, 78] = "Impreso";
                worksheet.Cells[Row, 79] = "Cancelado";
                worksheet.Cells[Row, 80] = "Total de unidades";
                worksheet.Cells[Row, 81] = "Clasificacion cliente2";
                worksheet.Cells[Row, 82] = "Texto extra1";
                worksheet.Cells[Row, 83] = "Nombre del concepto";

                Row++;
                #endregion

                #region contenido
                float total = 0;
                float pendiente = 0;
                for (int i = 0; i < ListDocmuentos.facturas.Count; i++)
                {
                    worksheet.Cells[Row, 1] = ListDocmuentos.facturas[i].Fecha;
                    worksheet.Cells[Row, 2] = ListDocmuentos.facturas[i].Referencia;
                    worksheet.Cells[Row, 3] = ListDocmuentos.facturas[i].Folio;
                    worksheet.Cells[Row, 4] = ListDocmuentos.facturas[i].NombreAgente;
                    worksheet.Cells[Row, 5] = ListDocmuentos.facturas[i].RazonSocial;
                    worksheet.Cells[Row, 6] = ListDocmuentos.facturas[i].FechaVencimiento;
                    worksheet.Cells[Row, 7] = ListDocmuentos.facturas[i].RFC;
                    worksheet.Cells[Row, 8] = ListDocmuentos.facturas[i].Subtotal;
                    worksheet.Cells[Row, 9] = ListDocmuentos.facturas[i].IVA;
                    worksheet.Cells[Row, 10] = ListDocmuentos.facturas[i].Total;
                    worksheet.Cells[Row, 11] = ListDocmuentos.facturas[i].Pendiente;
                    worksheet.Cells[Row, 12] = ListDocmuentos.facturas[i].TextoExtra3;
                    worksheet.Cells[Row, 13] = ListDocmuentos.facturas[i].Afectado;
                    worksheet.Cells[Row, 14] = ListDocmuentos.facturas[i].Impreso;
                    worksheet.Cells[Row, 15] = ListDocmuentos.facturas[i].Cancelado;
                    worksheet.Cells[Row, 16] = ListDocmuentos.facturas[i].TotalUnidades;
                    worksheet.Cells[Row, 17] = ListDocmuentos.facturas[i].Clasificacion2;
                    worksheet.Cells[Row, 18] = ListDocmuentos.facturas[i].TextoExtra1;
                    worksheet.Cells[Row, 19] = ListDocmuentos.facturas[i].NombreConcepto;
                    if (ListDocmuentos.facturas[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas[i].Total;
                        pendiente += ListDocmuentos.facturas[i].Pendiente;
                    }
                    Row++;
                }
                worksheet.Cells[2, 10] = "$ " + total;
                worksheet.Cells[2, 11] = "$ " + pendiente;

                total = 0;
                pendiente = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.facturas_dario.Count; i++)
                {
                    worksheet.Cells[Row, 23] = ListDocmuentos.facturas_dario[i].Fecha;
                    worksheet.Cells[Row, 24] = ListDocmuentos.facturas_dario[i].Referencia;
                    worksheet.Cells[Row, 25] = ListDocmuentos.facturas_dario[i].Folio;
                    worksheet.Cells[Row, 26] = ListDocmuentos.facturas_dario[i].NombreAgente;
                    worksheet.Cells[Row, 27] = ListDocmuentos.facturas_dario[i].RazonSocial;
                    worksheet.Cells[Row, 28] = ListDocmuentos.facturas_dario[i].FechaVencimiento;
                    worksheet.Cells[Row, 29] = ListDocmuentos.facturas_dario[i].RFC;
                    worksheet.Cells[Row, 30] = ListDocmuentos.facturas_dario[i].Subtotal;
                    worksheet.Cells[Row, 31] = ListDocmuentos.facturas_dario[i].IVA;
                    worksheet.Cells[Row, 32] = ListDocmuentos.facturas_dario[i].Total;
                    worksheet.Cells[Row, 33] = ListDocmuentos.facturas_dario[i].Pendiente;
                    worksheet.Cells[Row, 34] = ListDocmuentos.facturas_dario[i].TextoExtra3;
                    worksheet.Cells[Row, 35] = ListDocmuentos.facturas_dario[i].Afectado;
                    worksheet.Cells[Row, 36] = ListDocmuentos.facturas_dario[i].Impreso;
                    worksheet.Cells[Row, 37] = ListDocmuentos.facturas_dario[i].Cancelado;
                    worksheet.Cells[Row, 38] = ListDocmuentos.facturas_dario[i].TotalUnidades;
                    worksheet.Cells[Row, 39] = ListDocmuentos.facturas_dario[i].Clasificacion2;
                    worksheet.Cells[Row, 40] = ListDocmuentos.facturas_dario[i].TextoExtra1;
                    worksheet.Cells[Row, 41] = ListDocmuentos.facturas_dario[i].NombreConcepto;
                    if (ListDocmuentos.facturas_dario[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas_dario[i].Total;
                        pendiente += ListDocmuentos.facturas_dario[i].Pendiente;
                    }
                    Row++;
                }
                worksheet.Cells[2, 32] = "$ " + total;
                worksheet.Cells[2, 33] = "$ " + pendiente;


                total = 0;
                pendiente = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.abonos_dario.Count; i++)
                {
                    worksheet.Cells[Row, 65] = ListDocmuentos.abonos_dario[i].Fecha;
                    worksheet.Cells[Row, 66] = ListDocmuentos.abonos_dario[i].Referencia;
                    worksheet.Cells[Row, 67] = ListDocmuentos.abonos_dario[i].Folio;
                    worksheet.Cells[Row, 68] = ListDocmuentos.abonos_dario[i].NombreAgente;
                    worksheet.Cells[Row, 69] = ListDocmuentos.abonos_dario[i].RazonSocial;
                    worksheet.Cells[Row, 70] = ListDocmuentos.abonos_dario[i].FechaVencimiento;
                    worksheet.Cells[Row, 71] = ListDocmuentos.abonos_dario[i].RFC;
                    worksheet.Cells[Row, 72] = ListDocmuentos.abonos_dario[i].Subtotal;
                    worksheet.Cells[Row, 73] = ListDocmuentos.abonos_dario[i].IVA;
                    worksheet.Cells[Row, 74] = ListDocmuentos.abonos_dario[i].Total;
                    worksheet.Cells[Row, 75] = ListDocmuentos.abonos_dario[i].Pendiente;
                    worksheet.Cells[Row, 76] = ListDocmuentos.abonos_dario[i].TextoExtra3;
                    worksheet.Cells[Row, 77] = ListDocmuentos.abonos_dario[i].Afectado;
                    worksheet.Cells[Row, 78] = ListDocmuentos.abonos_dario[i].Impreso;
                    worksheet.Cells[Row, 79] = ListDocmuentos.abonos_dario[i].Cancelado;
                    worksheet.Cells[Row, 80] = ListDocmuentos.abonos_dario[i].TotalUnidades;
                    worksheet.Cells[Row, 81] = ListDocmuentos.abonos_dario[i].Clasificacion2;
                    worksheet.Cells[Row, 82] = ListDocmuentos.abonos_dario[i].TextoExtra1;
                    worksheet.Cells[Row, 83] = ListDocmuentos.abonos_dario[i].NombreConcepto;
                    total += ListDocmuentos.abonos_dario[i].Total;
                    pendiente += ListDocmuentos.abonos_dario[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 74] = "$ " + total;
                worksheet.Cells[2, 75] = "$ " + pendiente;



                total = 0;
                pendiente = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.abonos.Count; i++)
                {
                    worksheet.Cells[Row, 45] = ListDocmuentos.abonos[i].Fecha;
                    worksheet.Cells[Row, 46] = ListDocmuentos.abonos[i].Referencia;
                    worksheet.Cells[Row, 47] = ListDocmuentos.abonos[i].Folio;
                    worksheet.Cells[Row, 48] = ListDocmuentos.abonos[i].NombreAgente;
                    worksheet.Cells[Row, 49] = ListDocmuentos.abonos[i].RazonSocial;
                    worksheet.Cells[Row, 50] = ListDocmuentos.abonos[i].FechaVencimiento;
                    worksheet.Cells[Row, 51] = ListDocmuentos.abonos[i].RFC;
                    worksheet.Cells[Row, 52] = ListDocmuentos.abonos[i].Subtotal;
                    worksheet.Cells[Row, 53] = ListDocmuentos.abonos[i].IVA;
                    worksheet.Cells[Row, 54] = ListDocmuentos.abonos[i].Total;
                    worksheet.Cells[Row, 55] = ListDocmuentos.abonos[i].Pendiente;
                    worksheet.Cells[Row, 56] = ListDocmuentos.abonos[i].TextoExtra3;
                    worksheet.Cells[Row, 57] = ListDocmuentos.abonos[i].Afectado;
                    worksheet.Cells[Row, 58] = ListDocmuentos.abonos[i].Impreso;
                    worksheet.Cells[Row, 59] = ListDocmuentos.abonos[i].Cancelado;
                    worksheet.Cells[Row, 60] = ListDocmuentos.abonos[i].TotalUnidades;
                    worksheet.Cells[Row, 61] = ListDocmuentos.abonos[i].Clasificacion2;
                    worksheet.Cells[Row, 62] = ListDocmuentos.abonos[i].TextoExtra1;
                    worksheet.Cells[Row, 63] = ListDocmuentos.abonos[i].NombreConcepto;
                    total += ListDocmuentos.abonos[i].Total;
                    pendiente += ListDocmuentos.abonos[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 54] = "$ " + total;
                worksheet.Cells[2, 55] = "$ " + pendiente;

               

                #endregion



                #endregion




            }
            catch (Exception)
            {

            }
        }

        #endregion 

        #region IMPRT EXCEL Manuel
        public void excel_importMANUEL(Tipos_Datos_CRU.ListDatosMANUEL ListDocmuentos)
        { //importar datos en excel
            try
            {


                // creating Excel Application
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // see the excel sheet behind the program
                app.Visible = true;
                // get the reference of first sheet. By default its name is Sheet1.
                // store its reference to worksheet
                worksheet = workbook.Sheets["Hoja1"];
                worksheet = workbook.ActiveSheet;
                // changing the name of active sheet
                worksheet.Name = "Admipaq";

                #region configuracopn
                Microsoft.Office.Interop.Excel.Range formatRange;
                formatRange = worksheet.get_Range("A4", "s1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("V4", "AO1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                formatRange.Font.Size = 12;
                //
                formatRange = worksheet.get_Range("AS4", "VK1");
                formatRange.EntireRow.Font.Bold = true;
                formatRange.WrapText = true;//ajusta el texto a la columna
                formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                formatRange.Font.Size = 12;
                #endregion

                int Row = 4;
                //titulo
                #region todo
                Row = 4; //inicia a escribir en la fila 4
                #region encabezados
                worksheet.Cells[2, 5] = "facturas acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 1] = "Fecha";
                worksheet.Cells[Row, 2] = "Referencia";
                worksheet.Cells[Row, 3] = "Folio";
                worksheet.Cells[Row, 4] = "Nombre del agente";
                worksheet.Cells[Row, 5] = "Razon social";
                worksheet.Cells[Row, 6] = "Fecha de vencimiento";
                worksheet.Cells[Row, 7] = "RFC";
                worksheet.Cells[Row, 8] = "Subtotal";
                worksheet.Cells[Row, 9] = "IVA";
                worksheet.Cells[Row, 10] = "TOTAL";
                worksheet.Cells[Row, 11] = "Pendiente";
                worksheet.Cells[Row, 12] = "Texto Extra 3";
                worksheet.Cells[Row, 13] = "Afectado";
                worksheet.Cells[Row, 14] = "Impreso";
                worksheet.Cells[Row, 15] = "Cancelado";
                worksheet.Cells[Row, 16] = "Total de unidades";
                worksheet.Cells[Row, 17] = "Clasificacion cliente2";
                worksheet.Cells[Row, 18] = "Texto extra1";
                worksheet.Cells[Row, 19] = "Nombre del concepto";

                worksheet.Cells[2, 27] = "acumulado de compras";

                //encabezados facturas
                worksheet.Cells[Row, 22] = "Fecha";
                worksheet.Cells[Row, 24] = "Referencia";
                worksheet.Cells[Row, 25] = "Folio";
                worksheet.Cells[Row, 26] = "Nombre del agente";
                worksheet.Cells[Row, 27] = "Razon social";
                worksheet.Cells[Row, 28] = "Fecha de vencimiento";
                worksheet.Cells[Row, 29] = "RFC";
                worksheet.Cells[Row, 30] = "Subtotal";
                worksheet.Cells[Row, 31] = "IVA";
                worksheet.Cells[Row, 32] = "TOTAL";
                worksheet.Cells[Row, 33] = "Pendiente";
                worksheet.Cells[Row, 34] = "Texto Extra 3";
                worksheet.Cells[Row, 35] = "Afectado";
                worksheet.Cells[Row, 36] = "Impreso";
                worksheet.Cells[Row, 37] = "Cancelado";
                worksheet.Cells[Row, 38] = "Total de unidades";
                worksheet.Cells[Row, 39] = "Clasificacion cliente2";
                worksheet.Cells[Row, 40] = "Texto extra1";
                worksheet.Cells[Row, 41] = "Nombre del concepto";

                worksheet.Cells[2, 49] = "acumulado de pagos al proveedor";

                //encabezados facturas
                worksheet.Cells[Row, 45] = "Fecha";
                worksheet.Cells[Row, 46] = "Referencia";
                worksheet.Cells[Row, 47] = "Folio";
                worksheet.Cells[Row, 48] = "Nombre del agente";
                worksheet.Cells[Row, 49] = "Razon social";
                worksheet.Cells[Row, 50] = "Fecha de vencimiento";
                worksheet.Cells[Row, 51] = "RFC";
                worksheet.Cells[Row, 52] = "Subtotal";
                worksheet.Cells[Row, 53] = "IVA";
                worksheet.Cells[Row, 54] = "TOTAL";
                worksheet.Cells[Row, 55] = "Pendiente";
                worksheet.Cells[Row, 56] = "Texto Extra 3";
                worksheet.Cells[Row, 57] = "Afectado";
                worksheet.Cells[Row, 58] = "Impreso";
                worksheet.Cells[Row, 59] = "Cancelado";
                worksheet.Cells[Row, 60] = "Total de unidades";
                worksheet.Cells[Row, 61] = "Clasificacion cliente2";
                worksheet.Cells[Row, 62] = "Texto extra1";
                worksheet.Cells[Row, 63] = "Nombre del concepto";


                Row++;
                #endregion

                #region contenido
                float total = 0;
                float pendiente = 0;
                for (int i = 0; i < ListDocmuentos.facturas.Count; i++)
                {
                    worksheet.Cells[Row, 1] = ListDocmuentos.facturas[i].Fecha;
                    worksheet.Cells[Row, 2] = ListDocmuentos.facturas[i].Referencia;
                    worksheet.Cells[Row, 3] = ListDocmuentos.facturas[i].Folio;
                    worksheet.Cells[Row, 4] = ListDocmuentos.facturas[i].NombreAgente;
                    worksheet.Cells[Row, 5] = ListDocmuentos.facturas[i].RazonSocial;
                    worksheet.Cells[Row, 6] = ListDocmuentos.facturas[i].FechaVencimiento;
                    worksheet.Cells[Row, 7] = ListDocmuentos.facturas[i].RFC;
                    worksheet.Cells[Row, 8] = ListDocmuentos.facturas[i].Subtotal;
                    worksheet.Cells[Row, 9] = ListDocmuentos.facturas[i].IVA;
                    worksheet.Cells[Row, 10] = ListDocmuentos.facturas[i].Total;
                    worksheet.Cells[Row, 11] = ListDocmuentos.facturas[i].Pendiente;
                    worksheet.Cells[Row, 12] = ListDocmuentos.facturas[i].TextoExtra3;
                    worksheet.Cells[Row, 13] = ListDocmuentos.facturas[i].Afectado;
                    worksheet.Cells[Row, 14] = ListDocmuentos.facturas[i].Impreso;
                    worksheet.Cells[Row, 15] = ListDocmuentos.facturas[i].Cancelado;
                    worksheet.Cells[Row, 16] = ListDocmuentos.facturas[i].TotalUnidades;
                    worksheet.Cells[Row, 17] = ListDocmuentos.facturas[i].Clasificacion2;
                    worksheet.Cells[Row, 18] = ListDocmuentos.facturas[i].TextoExtra1;
                    worksheet.Cells[Row, 19] = ListDocmuentos.facturas[i].NombreConcepto;
                    if (ListDocmuentos.facturas[i].Cancelado.Trim() == "0")
                    {
                        total += ListDocmuentos.facturas[i].Total;
                        pendiente += ListDocmuentos.facturas[i].Pendiente;
                    }
                    Row++;
                }
                worksheet.Cells[2, 10] = "$ " + total;
                worksheet.Cells[2, 11] = "$ " + pendiente;

                total = 0;
                pendiente = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.compras.Count; i++)
                {
                    worksheet.Cells[Row, 23] = ListDocmuentos.compras[i].Fecha;
                    worksheet.Cells[Row, 24] = ListDocmuentos.compras[i].Referencia;
                    worksheet.Cells[Row, 25] = ListDocmuentos.compras[i].Folio;
                    worksheet.Cells[Row, 26] = ListDocmuentos.compras[i].NombreAgente;
                    worksheet.Cells[Row, 27] = ListDocmuentos.compras[i].RazonSocial;
                    worksheet.Cells[Row, 28] = ListDocmuentos.compras[i].FechaVencimiento;
                    worksheet.Cells[Row, 29] = ListDocmuentos.compras[i].RFC;
                    worksheet.Cells[Row, 30] = ListDocmuentos.compras[i].Subtotal;
                    worksheet.Cells[Row, 31] = ListDocmuentos.compras[i].IVA;
                    worksheet.Cells[Row, 32] = ListDocmuentos.compras[i].Total;
                    worksheet.Cells[Row, 33] = ListDocmuentos.compras[i].Pendiente;
                    worksheet.Cells[Row, 34] = ListDocmuentos.compras[i].TextoExtra3;
                    worksheet.Cells[Row, 35] = ListDocmuentos.compras[i].Afectado;
                    worksheet.Cells[Row, 36] = ListDocmuentos.compras[i].Impreso;
                    worksheet.Cells[Row, 37] = ListDocmuentos.compras[i].Cancelado;
                    worksheet.Cells[Row, 38] = ListDocmuentos.compras[i].TotalUnidades;
                    worksheet.Cells[Row, 39] = ListDocmuentos.compras[i].Clasificacion2;
                    worksheet.Cells[Row, 40] = ListDocmuentos.compras[i].TextoExtra1;
                    worksheet.Cells[Row, 41] = ListDocmuentos.compras[i].NombreConcepto;
                    total += ListDocmuentos.compras[i].Total;
                    pendiente += ListDocmuentos.compras[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 32] = "$ " + total;
                worksheet.Cells[2, 33] = "$ " + pendiente;

                total = 0;
                pendiente = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.pagosproveedor.Count; i++)
                {
                    worksheet.Cells[Row, 45] = ListDocmuentos.pagosproveedor[i].Fecha;
                    worksheet.Cells[Row, 46] = ListDocmuentos.pagosproveedor[i].Referencia;
                    worksheet.Cells[Row, 47] = ListDocmuentos.pagosproveedor[i].Folio;
                    worksheet.Cells[Row, 48] = ListDocmuentos.pagosproveedor[i].NombreAgente;
                    worksheet.Cells[Row, 49] = ListDocmuentos.pagosproveedor[i].RazonSocial;
                    worksheet.Cells[Row, 50] = ListDocmuentos.pagosproveedor[i].FechaVencimiento;
                    worksheet.Cells[Row, 51] = ListDocmuentos.pagosproveedor[i].RFC;
                    worksheet.Cells[Row, 52] = ListDocmuentos.pagosproveedor[i].Subtotal;
                    worksheet.Cells[Row, 53] = ListDocmuentos.pagosproveedor[i].IVA;
                    worksheet.Cells[Row, 54] = ListDocmuentos.pagosproveedor[i].Total;
                    worksheet.Cells[Row, 55] = ListDocmuentos.pagosproveedor[i].Pendiente;
                    worksheet.Cells[Row, 56] = ListDocmuentos.pagosproveedor[i].TextoExtra3;
                    worksheet.Cells[Row, 57] = ListDocmuentos.pagosproveedor[i].Afectado;
                    worksheet.Cells[Row, 58] = ListDocmuentos.pagosproveedor[i].Impreso;
                    worksheet.Cells[Row, 59] = ListDocmuentos.pagosproveedor[i].Cancelado;
                    worksheet.Cells[Row, 60] = ListDocmuentos.pagosproveedor[i].TotalUnidades;
                    worksheet.Cells[Row, 61] = ListDocmuentos.pagosproveedor[i].Clasificacion2;
                    worksheet.Cells[Row, 62] = ListDocmuentos.pagosproveedor[i].TextoExtra1;
                    worksheet.Cells[Row, 63] = ListDocmuentos.pagosproveedor[i].NombreConcepto;
                    total += ListDocmuentos.pagosproveedor[i].Total;
                    pendiente += ListDocmuentos.pagosproveedor[i].Pendiente;
                    Row++;
                }
                worksheet.Cells[2, 54] = "$ " + total;
                worksheet.Cells[2, 55] = "$ " + pendiente;

                

                #endregion



                #endregion


                


            }
            catch (Exception)
            {

            }
        }

        #endregion 
    }
}
