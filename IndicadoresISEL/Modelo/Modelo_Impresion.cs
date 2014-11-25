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
                    total += ListDocmuentos.facturas[i].Total;
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
                    total += ListDocmuentos.facturas_rfc_publico[i].Total;
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
                    total += ListDocmuentos.facturas_rfc_ol[i].Total;
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
                worksheet.Cells[2, 112] = "Abonos RFC";
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
                worksheet.Cells[Row, 125] = "Texto extra1";
                worksheet.Cells[Row, 126] = "Nombre del concepto";

                worksheet.Cells[2, 133] = "Zonas centro";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 129] = "Fecha";
                worksheet.Cells[Row, 130] = "Serie";
                worksheet.Cells[Row, 131] = "Folio";
                worksheet.Cells[Row, 132] = "Nombre del agente";
                worksheet.Cells[Row, 133] = "Razon social";
                worksheet.Cells[Row, 134] = "Fecha de vencimiento";
                worksheet.Cells[Row, 135] = "RFC";
                worksheet.Cells[Row, 136] = "Subtotal";
                worksheet.Cells[Row, 137] = "IVA";
                worksheet.Cells[Row, 138] = "TOTAL";
                worksheet.Cells[Row, 139] = "Pendiente";
                worksheet.Cells[Row, 140] = "Texto Extra 3";
                worksheet.Cells[Row, 141] = "Afectado";
                worksheet.Cells[Row, 142] = "Impreso";
                worksheet.Cells[Row, 143] = "Cancelado";
                worksheet.Cells[Row, 144] = "Total de unidades";
                worksheet.Cells[Row, 145] = "Zona ";
                worksheet.Cells[Row, 146] = "Agente";

                worksheet.Cells[2, 153] = "Zonas sur";

                worksheet.Cells[Row, 149] = "Fecha";
                worksheet.Cells[Row, 150] = "Serie";
                worksheet.Cells[Row, 151] = "Folio";
                worksheet.Cells[Row, 152] = "Nombre del agente";
                worksheet.Cells[Row, 153] = "Razon social";
                worksheet.Cells[Row, 154] = "Fecha de vencimiento";
                worksheet.Cells[Row, 155] = "RFC";
                worksheet.Cells[Row, 156] = "Subtotal";
                worksheet.Cells[Row, 157] = "IVA";
                worksheet.Cells[Row, 158] = "TOTAL";
                worksheet.Cells[Row, 159] = "Pendiente";
                worksheet.Cells[Row, 160] = "Texto Extra 3";
                worksheet.Cells[Row, 161] = "Afectado";
                worksheet.Cells[Row, 162] = "Impreso";
                worksheet.Cells[Row, 163] = "Cancelado";
                worksheet.Cells[Row, 164] = "Total de unidades";
                worksheet.Cells[Row, 165] = "Zona";
                worksheet.Cells[Row, 166] = "Agente";

                worksheet.Cells[2, 173] = "Zonas norte";

                worksheet.Cells[Row, 169] = "Fecha";
                worksheet.Cells[Row, 170] = "Serie";
                worksheet.Cells[Row, 171] = "Folio";
                worksheet.Cells[Row, 172] = "Nombre del agente";
                worksheet.Cells[Row, 173] = "Razon social";
                worksheet.Cells[Row, 174] = "Fecha de vencimiento";
                worksheet.Cells[Row, 175] = "RFC";
                worksheet.Cells[Row, 176] = "Subtotal";
                worksheet.Cells[Row, 177] = "IVA";
                worksheet.Cells[Row, 178] = "TOTAL";
                worksheet.Cells[Row, 179] = "Pendiente";
                worksheet.Cells[Row, 180] = "Texto Extra 3";
                worksheet.Cells[Row, 181] = "Afectado";
                worksheet.Cells[Row, 182] = "Impreso";
                worksheet.Cells[Row, 183] = "Cancelado";
                worksheet.Cells[Row, 184] = "Total de unidades";
                worksheet.Cells[Row, 185] = "Zona";
                worksheet.Cells[Row, 186] = "Agente";
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
                worksheet.Cells[2, 76] = "$ " + total;

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
                    worksheet.Cells[Row, 125] = ListDocmuentos.abonos_ol[i].TextoExtra1;
                    worksheet.Cells[Row, 126] = ListDocmuentos.abonos_ol[i].NombreConcepto;
                    total += ListDocmuentos.abonos_ol[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 117] = "$ " + total;
                total = 0;
                Row = 5;


                /*ZONA CENTRO*/
                for (int i = 0; i < ListDocmuentos.abonos_zona_norte.Count; i++)
                {
                    worksheet.Cells[Row, 129] = ListDocmuentos.abonos_zona_norte[i].Fecha;
                    worksheet.Cells[Row, 130] = ListDocmuentos.abonos_zona_norte[i].Serie;
                    worksheet.Cells[Row, 131] = ListDocmuentos.abonos_zona_norte[i].Folio;
                    worksheet.Cells[Row, 132] = ListDocmuentos.abonos_zona_norte[i].NombreAgente;
                    worksheet.Cells[Row, 133] = ListDocmuentos.abonos_zona_norte[i].RazonSocial;
                    worksheet.Cells[Row, 134] = ListDocmuentos.abonos_zona_norte[i].FechaVencimiento;
                    worksheet.Cells[Row, 135] = ListDocmuentos.abonos_zona_norte[i].RFC;
                    worksheet.Cells[Row, 136] = ListDocmuentos.abonos_zona_norte[i].Subtotal;
                    worksheet.Cells[Row, 137] = ListDocmuentos.abonos_zona_norte[i].IVA;
                    worksheet.Cells[Row, 138] = ListDocmuentos.abonos_zona_norte[i].Total;
                    worksheet.Cells[Row, 139] = ListDocmuentos.abonos_zona_norte[i].Pendiente;
                    worksheet.Cells[Row, 140] = ListDocmuentos.abonos_zona_norte[i].TextoExtra3;
                    worksheet.Cells[Row, 141] = ListDocmuentos.abonos_zona_norte[i].Afectado;
                    worksheet.Cells[Row, 142] = ListDocmuentos.abonos_zona_norte[i].Impreso;
                    worksheet.Cells[Row, 143] = ListDocmuentos.abonos_zona_norte[i].Cancelado;
                    worksheet.Cells[Row, 144] = ListDocmuentos.abonos_zona_norte[i].TotalUnidades;
                    worksheet.Cells[Row, 145] = ListDocmuentos.abonos_zona_norte[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 146] = ListDocmuentos.abonos_zona_norte[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_norte[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 138] = "$ " + total;
                total = 0;
                Row = 5;
                /*ZONA sur*/
                for (int i = 0; i < ListDocmuentos.abonos_zona_sur.Count; i++)
                {
                    worksheet.Cells[Row, 149] = ListDocmuentos.abonos_zona_sur[i].Fecha;
                    worksheet.Cells[Row, 150] = ListDocmuentos.abonos_zona_sur[i].Serie;
                    worksheet.Cells[Row, 151] = ListDocmuentos.abonos_zona_sur[i].Folio;
                    worksheet.Cells[Row, 152] = ListDocmuentos.abonos_zona_sur[i].NombreAgente;
                    worksheet.Cells[Row, 153] = ListDocmuentos.abonos_zona_sur[i].RazonSocial;
                    worksheet.Cells[Row, 154] = ListDocmuentos.abonos_zona_sur[i].FechaVencimiento;
                    worksheet.Cells[Row, 155] = ListDocmuentos.abonos_zona_sur[i].RFC;
                    worksheet.Cells[Row, 156] = ListDocmuentos.abonos_zona_sur[i].Subtotal;
                    worksheet.Cells[Row, 157] = ListDocmuentos.abonos_zona_sur[i].IVA;
                    worksheet.Cells[Row, 158] = ListDocmuentos.abonos_zona_sur[i].Total;
                    worksheet.Cells[Row, 159] = ListDocmuentos.abonos_zona_sur[i].Pendiente;
                    worksheet.Cells[Row, 160] = ListDocmuentos.abonos_zona_sur[i].TextoExtra3;
                    worksheet.Cells[Row, 161] = ListDocmuentos.abonos_zona_sur[i].Afectado;
                    worksheet.Cells[Row, 162] = ListDocmuentos.abonos_zona_sur[i].Impreso;
                    worksheet.Cells[Row, 163] = ListDocmuentos.abonos_zona_sur[i].Cancelado;
                    worksheet.Cells[Row, 164] = ListDocmuentos.abonos_zona_sur[i].TotalUnidades;
                    worksheet.Cells[Row, 165] = ListDocmuentos.abonos_zona_sur[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 166] = ListDocmuentos.abonos_zona_sur[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_sur[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 158] = "$ " + total;
                total = 0;
                Row = 5;
                /*ZONA norte*/
                for (int i = 0; i < ListDocmuentos.abonos_zona_norte.Count; i++)
                {
                    worksheet.Cells[Row, 169] = ListDocmuentos.abonos_zona_norte[i].Fecha;
                    worksheet.Cells[Row, 170] = ListDocmuentos.abonos_zona_norte[i].Serie;
                    worksheet.Cells[Row, 171] = ListDocmuentos.abonos_zona_norte[i].Folio;
                    worksheet.Cells[Row, 172] = ListDocmuentos.abonos_zona_norte[i].NombreAgente;
                    worksheet.Cells[Row, 173] = ListDocmuentos.abonos_zona_norte[i].RazonSocial;
                    worksheet.Cells[Row, 174] = ListDocmuentos.abonos_zona_norte[i].FechaVencimiento;
                    worksheet.Cells[Row, 175] = ListDocmuentos.abonos_zona_norte[i].RFC;
                    worksheet.Cells[Row, 176] = ListDocmuentos.abonos_zona_norte[i].Subtotal;
                    worksheet.Cells[Row, 177] = ListDocmuentos.abonos_zona_norte[i].IVA;
                    worksheet.Cells[Row, 178] = ListDocmuentos.abonos_zona_norte[i].Total;
                    worksheet.Cells[Row, 179] = ListDocmuentos.abonos_zona_norte[i].Pendiente;
                    worksheet.Cells[Row, 180] = ListDocmuentos.abonos_zona_norte[i].TextoExtra3;
                    worksheet.Cells[Row, 181] = ListDocmuentos.abonos_zona_norte[i].Afectado;
                    worksheet.Cells[Row, 182] = ListDocmuentos.abonos_zona_norte[i].Impreso;
                    worksheet.Cells[Row, 183] = ListDocmuentos.abonos_zona_norte[i].Cancelado;
                    worksheet.Cells[Row, 184] = ListDocmuentos.abonos_zona_norte[i].TotalUnidades;
                    worksheet.Cells[Row, 185] = ListDocmuentos.abonos_zona_norte[i].proveedor.Clasificación1;
                    worksheet.Cells[Row, 186] = ListDocmuentos.abonos_zona_norte[i].proveedor.Clasificación2;
                    total += ListDocmuentos.abonos_zona_norte[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 178] = "$ " + total;
                #endregion
                #endregion

                #region Compras
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 194] = "compras acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 190] = "Fecha";
                worksheet.Cells[Row, 191] = "Serie";
                worksheet.Cells[Row, 192] = "Folio";
                worksheet.Cells[Row, 193] = "Nombre del agente";
                worksheet.Cells[Row, 194] = "Razon social";
                worksheet.Cells[Row, 195] = "Fecha de vencimiento";
                worksheet.Cells[Row, 196] = "RFC";
                worksheet.Cells[Row, 197] = "Subtotal";
                worksheet.Cells[Row,198] = "IVA";
                worksheet.Cells[Row, 199] = "TOTAL";
                worksheet.Cells[Row, 200] = "Pendiente";
                worksheet.Cells[Row, 201] = "Texto Extra 3";
                worksheet.Cells[Row, 202] = "Afectado";
                worksheet.Cells[Row, 203] = "Impreso";
                worksheet.Cells[Row, 204] = "Cancelado";
                worksheet.Cells[Row, 205] = "Total de unidades";
                worksheet.Cells[Row, 206] = "Clasificacion cliente2";
                worksheet.Cells[Row, 207] = "Texto extra1";
                worksheet.Cells[Row, 208] = "Nombre del concepto";
                //titulo 
                worksheet.Cells[2, 215] = "compras ANJI";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 211] = "Fecha";
                worksheet.Cells[Row, 212] = "Serie";
                worksheet.Cells[Row, 213] = "Folio";
                worksheet.Cells[Row, 214] = "Nombre del agente";
                worksheet.Cells[Row, 215] = "Razon social";
                worksheet.Cells[Row, 216] = "Fecha de vencimiento";
                worksheet.Cells[Row, 217] = "RFC";
                worksheet.Cells[Row, 218] = "Subtotal";
                worksheet.Cells[Row, 219] = "IVA";
                worksheet.Cells[Row, 220] = "TOTAL";
                worksheet.Cells[Row, 221] = "Pendiente";
                worksheet.Cells[Row, 222] = "Texto Extra 3";
                worksheet.Cells[Row, 223] = "Afectado";
                worksheet.Cells[Row, 224] = "Impreso";
                worksheet.Cells[Row, 225] = "Cancelado";
                worksheet.Cells[Row, 226] = "Total de unidades";
                worksheet.Cells[Row, 227] = "Clasificacion cliente2";
                worksheet.Cells[Row, 228] = "Texto extra1";
                worksheet.Cells[Row, 229] = "Nombre del concepto";

                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.compras.Count; i++)
                {
                    worksheet.Cells[Row, 190] = ListDocmuentos.compras[i].Fecha;
                    worksheet.Cells[Row, 191] = ListDocmuentos.compras[i].Serie;
                    worksheet.Cells[Row, 192] = ListDocmuentos.compras[i].Folio;
                    worksheet.Cells[Row, 193] = ListDocmuentos.compras[i].NombreAgente;
                    worksheet.Cells[Row, 194] = ListDocmuentos.compras[i].RazonSocial;
                    worksheet.Cells[Row, 195] = ListDocmuentos.compras[i].FechaVencimiento;
                    worksheet.Cells[Row, 196] = ListDocmuentos.compras[i].RFC;
                    worksheet.Cells[Row, 197] = ListDocmuentos.compras[i].Subtotal;
                    worksheet.Cells[Row, 198] = ListDocmuentos.compras[i].IVA;
                    worksheet.Cells[Row, 199] = ListDocmuentos.compras[i].Total;
                    worksheet.Cells[Row, 200] = ListDocmuentos.compras[i].Pendiente;
                    worksheet.Cells[Row, 201] = ListDocmuentos.compras[i].TextoExtra3;
                    worksheet.Cells[Row, 202] = ListDocmuentos.compras[i].Afectado;
                    worksheet.Cells[Row, 203] = ListDocmuentos.compras[i].Impreso;
                    worksheet.Cells[Row, 204] = ListDocmuentos.compras[i].Cancelado;
                    worksheet.Cells[Row, 205] = ListDocmuentos.compras[i].TotalUnidades;
                    worksheet.Cells[Row, 206] = ListDocmuentos.compras[i].Clasificacion2;
                    worksheet.Cells[Row, 207] = ListDocmuentos.compras[i].TextoExtra1;
                    worksheet.Cells[Row, 208] = ListDocmuentos.compras[i].NombreConcepto;
                    total += ListDocmuentos.compras[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 199] = "$ " + total;

                total = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.compras_rfc_anji.Count; i++)
                {
                    worksheet.Cells[Row, 211] = ListDocmuentos.compras_rfc_anji[i].Fecha;
                    worksheet.Cells[Row, 212] = ListDocmuentos.compras_rfc_anji[i].Serie;
                    worksheet.Cells[Row, 213] = ListDocmuentos.compras_rfc_anji[i].Folio;
                    worksheet.Cells[Row, 214] = ListDocmuentos.compras_rfc_anji[i].NombreAgente;
                    worksheet.Cells[Row, 215] = ListDocmuentos.compras_rfc_anji[i].RazonSocial;
                    worksheet.Cells[Row, 216] = ListDocmuentos.compras_rfc_anji[i].FechaVencimiento;
                    worksheet.Cells[Row, 217] = ListDocmuentos.compras_rfc_anji[i].RFC;
                    worksheet.Cells[Row, 218] = ListDocmuentos.compras_rfc_anji[i].Subtotal;
                    worksheet.Cells[Row, 219] = ListDocmuentos.compras_rfc_anji[i].IVA;
                    worksheet.Cells[Row, 220] = ListDocmuentos.compras_rfc_anji[i].Total;
                    worksheet.Cells[Row, 221] = ListDocmuentos.compras_rfc_anji[i].Pendiente;
                    worksheet.Cells[Row, 222] = ListDocmuentos.compras_rfc_anji[i].TextoExtra3;
                    worksheet.Cells[Row, 223] = ListDocmuentos.compras_rfc_anji[i].Afectado;
                    worksheet.Cells[Row, 224] = ListDocmuentos.compras_rfc_anji[i].Impreso;
                    worksheet.Cells[Row, 225] = ListDocmuentos.compras_rfc_anji[i].Cancelado;
                    worksheet.Cells[Row, 226] = ListDocmuentos.compras_rfc_anji[i].TotalUnidades;
                    worksheet.Cells[Row, 227] = ListDocmuentos.compras_rfc_anji[i].Clasificacion2;
                    worksheet.Cells[Row, 228] = ListDocmuentos.compras_rfc_anji[i].TextoExtra1;
                    worksheet.Cells[Row, 229] = ListDocmuentos.compras_rfc_anji[i].NombreConcepto;
                    total += ListDocmuentos.compras_rfc_anji[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 220] = "$ " + total;
                #endregion
                #endregion

                #region Pagos proveedor
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 238] = "pagos acumuladas";

                //encabezados facturas
                worksheet.Cells[Row, 234] = "Fecha";
                worksheet.Cells[Row, 235] = "Serie";
                worksheet.Cells[Row, 236] = "Folio";
                worksheet.Cells[Row, 237] = "Nombre del agente";
                worksheet.Cells[Row, 238] = "Razon social";
                worksheet.Cells[Row, 239] = "Fecha de vencimiento";
                worksheet.Cells[Row, 240] = "RFC";
                worksheet.Cells[Row, 241] = "Subtotal";
                worksheet.Cells[Row, 242] = "IVA";
                worksheet.Cells[Row, 243] = "TOTAL";
                worksheet.Cells[Row, 244] = "Pendiente";
                worksheet.Cells[Row, 245] = "Texto Extra 3";
                worksheet.Cells[Row, 246] = "Afectado";
                worksheet.Cells[Row, 247] = "Impreso";
                worksheet.Cells[Row, 248] = "Cancelado";
                worksheet.Cells[Row, 249] = "Total de unidades";
                worksheet.Cells[Row, 250] = "Clasificacion cliente2";
                worksheet.Cells[Row, 251] = "Texto extra1";
                worksheet.Cells[Row, 252] = "Nombre del concepto";
                //titulo 
                worksheet.Cells[2, 259] = "pagos Anji";
                //envabezados facturas filtro publico
                worksheet.Cells[Row, 255] = "Fecha";
                worksheet.Cells[Row, 256] = "Serie";
                worksheet.Cells[Row, 257] = "Folio";
                worksheet.Cells[Row, 258] = "Nombre del agente";
                worksheet.Cells[Row, 259] = "Razon social";
                worksheet.Cells[Row, 260] = "Fecha de vencimiento";
                worksheet.Cells[Row, 261] = "RFC";
                worksheet.Cells[Row, 262] = "Subtotal";
                worksheet.Cells[Row, 263] = "IVA";
                worksheet.Cells[Row, 264] = "TOTAL";
                worksheet.Cells[Row, 265] = "Pendiente";
                worksheet.Cells[Row, 266] = "Texto Extra 3";
                worksheet.Cells[Row, 267] = "Afectado";
                worksheet.Cells[Row, 268] = "Impreso";
                worksheet.Cells[Row, 269] = "Cancelado";
                worksheet.Cells[Row, 270] = "Total de unidades";
                worksheet.Cells[Row, 271] = "Clasificacion cliente2";
                worksheet.Cells[Row, 272] = "Texto extra1";
                worksheet.Cells[Row, 273] = "Nombre del concepto";
                
                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.pagos_proveedor.Count; i++)
                {
                    worksheet.Cells[Row, 234] = ListDocmuentos.pagos_proveedor[i].Fecha;
                    worksheet.Cells[Row, 235] = ListDocmuentos.pagos_proveedor[i].Serie;
                    worksheet.Cells[Row, 236] = ListDocmuentos.pagos_proveedor[i].Folio;
                    worksheet.Cells[Row, 237] = ListDocmuentos.pagos_proveedor[i].NombreAgente;
                    worksheet.Cells[Row, 238] = ListDocmuentos.pagos_proveedor[i].RazonSocial;
                    worksheet.Cells[Row, 239] = ListDocmuentos.pagos_proveedor[i].FechaVencimiento;
                    worksheet.Cells[Row, 240] = ListDocmuentos.pagos_proveedor[i].RFC;
                    worksheet.Cells[Row, 241] = ListDocmuentos.pagos_proveedor[i].Subtotal;
                    worksheet.Cells[Row, 242] = ListDocmuentos.pagos_proveedor[i].IVA;
                    worksheet.Cells[Row, 243] = ListDocmuentos.pagos_proveedor[i].Total;
                    worksheet.Cells[Row, 244] = ListDocmuentos.pagos_proveedor[i].Pendiente;
                    worksheet.Cells[Row, 245] = ListDocmuentos.pagos_proveedor[i].TextoExtra3;
                    worksheet.Cells[Row, 246] = ListDocmuentos.pagos_proveedor[i].Afectado;
                    worksheet.Cells[Row, 247] = ListDocmuentos.pagos_proveedor[i].Impreso;
                    worksheet.Cells[Row, 248] = ListDocmuentos.pagos_proveedor[i].Cancelado;
                    worksheet.Cells[Row, 249] = ListDocmuentos.pagos_proveedor[i].TotalUnidades;
                    worksheet.Cells[Row, 250] = ListDocmuentos.pagos_proveedor[i].Clasificacion2;
                    worksheet.Cells[Row, 251] = ListDocmuentos.pagos_proveedor[i].TextoExtra1;
                    worksheet.Cells[Row, 252] = ListDocmuentos.pagos_proveedor[i].NombreConcepto;
                    total += ListDocmuentos.pagos_proveedor[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 10] = "$ " + total;

                total = 0;
                Row = 5;

                for (int i = 0; i < ListDocmuentos.pagos_proveedor_rfc_anji.Count; i++)
                {
                    worksheet.Cells[Row, 255] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Fecha;
                    worksheet.Cells[Row, 256] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Serie;
                    worksheet.Cells[Row, 257] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Folio;
                    worksheet.Cells[Row, 258] = ListDocmuentos.pagos_proveedor_rfc_anji[i].NombreAgente;
                    worksheet.Cells[Row, 259] = ListDocmuentos.pagos_proveedor_rfc_anji[i].RazonSocial;
                    worksheet.Cells[Row, 260] = ListDocmuentos.pagos_proveedor_rfc_anji[i].FechaVencimiento;
                    worksheet.Cells[Row, 261] = ListDocmuentos.pagos_proveedor_rfc_anji[i].RFC;
                    worksheet.Cells[Row, 262] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Subtotal;
                    worksheet.Cells[Row, 263] = ListDocmuentos.pagos_proveedor_rfc_anji[i].IVA;
                    worksheet.Cells[Row, 264] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Total;
                    worksheet.Cells[Row, 265] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Pendiente;
                    worksheet.Cells[Row, 266] = ListDocmuentos.pagos_proveedor_rfc_anji[i].TextoExtra3;
                    worksheet.Cells[Row, 267] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Afectado;
                    worksheet.Cells[Row, 268] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Impreso;
                    worksheet.Cells[Row, 269] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Cancelado;
                    worksheet.Cells[Row, 270] = ListDocmuentos.pagos_proveedor_rfc_anji[i].TotalUnidades;
                    worksheet.Cells[Row, 271] = ListDocmuentos.pagos_proveedor_rfc_anji[i].Clasificacion2;
                    worksheet.Cells[Row, 272] = ListDocmuentos.pagos_proveedor_rfc_anji[i].TextoExtra1;
                    worksheet.Cells[Row, 273] = ListDocmuentos.pagos_proveedor_rfc_anji[i].NombreConcepto;
                    total += ListDocmuentos.pagos_proveedor_rfc_anji[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 264] = "$ " + total;
                
                #endregion
                #endregion

                #region Prestamos
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 277] = "Prestamos";

                //encabezados prestamos
                worksheet.Cells[Row, 277] = "Fecha";
                worksheet.Cells[Row, 274] = "Serie";
                worksheet.Cells[Row, 275] = "Folio";
                worksheet.Cells[Row, 276] = "Nombre del agente";
                worksheet.Cells[Row, 277] = "Razon social";
                worksheet.Cells[Row, 278] = "Fecha de vencimiento";
                worksheet.Cells[Row, 279] = "Fecha de depósito";
                worksheet.Cells[Row, 280] = "RFC";
                worksheet.Cells[Row, 281] = "TOTAL";
                worksheet.Cells[Row, 282] = "Pendiente";
                worksheet.Cells[Row, 283] = "Cuenta";
                worksheet.Cells[Row, 284] = "Referencia";

                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.prestamos.Count; i++)
                {
                    worksheet.Cells[Row, 277] = ListDocmuentos.prestamos[i].Fecha;
                    worksheet.Cells[Row, 278] = ListDocmuentos.prestamos[i].Serie;
                    worksheet.Cells[Row, 279] = ListDocmuentos.prestamos[i].Folio;
                    worksheet.Cells[Row, 280] = ListDocmuentos.prestamos[i].NombreAgente;
                    worksheet.Cells[Row, 281] = ListDocmuentos.prestamos[i].RazonSocial;
                    worksheet.Cells[Row, 282] = ListDocmuentos.prestamos[i].FechaVencimiento;
                    worksheet.Cells[Row, 283] = ListDocmuentos.prestamos[i].TextoExtra1;
                    worksheet.Cells[Row, 284] = ListDocmuentos.prestamos[i].RFC;
                    worksheet.Cells[Row, 285] = ListDocmuentos.prestamos[i].Total;
                    worksheet.Cells[Row, 286] = ListDocmuentos.prestamos[i].Pendiente;
                    worksheet.Cells[Row, 287] = ListDocmuentos.prestamos[i].TextoExtra2;
                    worksheet.Cells[Row, 288] = ListDocmuentos.prestamos[i].Referencia;
                    total += ListDocmuentos.prestamos[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 285] = "$ " + total;
                #endregion
                #endregion

                #region Ingreso traspaso
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 294] = "Ingreso traspaso";

                //encabezados ingtreso traspaso
                worksheet.Cells[Row, 290] = "Fecha";
                worksheet.Cells[Row, 291] = "Serie";
                worksheet.Cells[Row, 292] = "Folio";
                worksheet.Cells[Row, 293] = "Nombre del agente";
                worksheet.Cells[Row, 294] = "Razon social";
                worksheet.Cells[Row, 295] = "Fecha de vencimiento";
                worksheet.Cells[Row, 296] = "Fecha de depósito";
                worksheet.Cells[Row, 297] = "RFC";
                worksheet.Cells[Row, 298] = "TOTAL";
                worksheet.Cells[Row, 299] = "Pendiente";
                worksheet.Cells[Row, 300] = "texto extra 2";
                worksheet.Cells[Row, 301] = "Referencia";

                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.ingreso_traspaso.Count; i++)
                {
                    worksheet.Cells[Row, 290] = ListDocmuentos.ingreso_traspaso[i].Fecha;
                    worksheet.Cells[Row, 291] = ListDocmuentos.ingreso_traspaso[i].Serie;
                    worksheet.Cells[Row, 292] = ListDocmuentos.ingreso_traspaso[i].Folio;
                    worksheet.Cells[Row, 293] = ListDocmuentos.ingreso_traspaso[i].NombreAgente;
                    worksheet.Cells[Row, 294] = ListDocmuentos.ingreso_traspaso[i].RazonSocial;
                    worksheet.Cells[Row, 295] = ListDocmuentos.ingreso_traspaso[i].FechaVencimiento;
                    worksheet.Cells[Row, 296] = ListDocmuentos.ingreso_traspaso[i].TextoExtra1;
                    worksheet.Cells[Row, 297] = ListDocmuentos.ingreso_traspaso[i].RFC;
                    worksheet.Cells[Row, 298] = ListDocmuentos.ingreso_traspaso[i].Total;
                    worksheet.Cells[Row, 299] = ListDocmuentos.ingreso_traspaso[i].Pendiente;
                    worksheet.Cells[Row, 300] = ListDocmuentos.ingreso_traspaso[i].TextoExtra2;
                    worksheet.Cells[Row, 301] = ListDocmuentos.ingreso_traspaso[i].Referencia;
                    total += ListDocmuentos.ingreso_traspaso[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 298] = "$ " + total;
                #endregion
                #endregion

                #region Ingreso dev. garantia
                Row = 4;
                #region encabezados
                worksheet.Cells[2, 309] = "Ingreso Dev. garantía";

                //encabezados ingtreso traspaso
                worksheet.Cells[Row, 305] = "Fecha";
                worksheet.Cells[Row, 306] = "Serie";
                worksheet.Cells[Row, 307] = "Folio";
                worksheet.Cells[Row, 308] = "Nombre del agente";
                worksheet.Cells[Row, 309] = "Razon social";
                worksheet.Cells[Row, 310] = "Fecha de vencimiento";
                worksheet.Cells[Row, 311] = "Fecha de depósito";
                worksheet.Cells[Row, 312] = "RFC";
                worksheet.Cells[Row, 313] = "TOTAL";
                worksheet.Cells[Row, 314] = "Pendiente";
                worksheet.Cells[Row, 315] = "texto extra 2";
                worksheet.Cells[Row, 316] = "Referencia";

                Row++;
                #endregion
                #region contenido
                total = 0;
                for (int i = 0; i < ListDocmuentos.ingreso_dev_garantia.Count; i++)
                {
                    worksheet.Cells[Row, 305] = ListDocmuentos.ingreso_dev_garantia[i].Fecha;
                    worksheet.Cells[Row, 306] = ListDocmuentos.ingreso_dev_garantia[i].Serie;
                    worksheet.Cells[Row, 307] = ListDocmuentos.ingreso_dev_garantia[i].Folio;
                    worksheet.Cells[Row, 308] = ListDocmuentos.ingreso_dev_garantia[i].NombreAgente;
                    worksheet.Cells[Row, 309] = ListDocmuentos.ingreso_dev_garantia[i].RazonSocial;
                    worksheet.Cells[Row, 310] = ListDocmuentos.ingreso_dev_garantia[i].FechaVencimiento;
                    worksheet.Cells[Row, 311] = ListDocmuentos.ingreso_dev_garantia[i].TextoExtra1;
                    worksheet.Cells[Row, 312] = ListDocmuentos.ingreso_dev_garantia[i].RFC;
                    worksheet.Cells[Row, 313] = ListDocmuentos.ingreso_dev_garantia[i].Total;
                    worksheet.Cells[Row, 314] = ListDocmuentos.ingreso_dev_garantia[i].Pendiente;
                    worksheet.Cells[Row, 315] = ListDocmuentos.ingreso_dev_garantia[i].TextoExtra2;
                    worksheet.Cells[Row, 316] = ListDocmuentos.ingreso_dev_garantia[i].Referencia;
                    total += ListDocmuentos.ingreso_dev_garantia[i].Total;
                    Row++;
                }
                worksheet.Cells[2, 313] = "$ " + total;
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
