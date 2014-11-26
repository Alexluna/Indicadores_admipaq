using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IndicadoresISEL.Modelo;
using System.Data;
using System.Windows;
using System.IO;
using System.Windows.Controls;

namespace IndicadoresISEL.Controlador
{
    class Controlador__SDKAdmipaq
    {
        Modelo_SDKAdmipaq ModeloSDK;//objeto para comunicarse con el modelos del sdk
        bool Conexion;//variable que me dira si tengo conexion con el sdk de admipaq
        int val_Cargador;//valor que tendra el cargador
        /// <summary>
        /// Constructor 
        /// </summary>
        public Controlador__SDKAdmipaq()
        {
            ModeloSDK = new Modelo_SDKAdmipaq();
            Conexion = false;//inicializo sin conexion
            val_Cargador = 0;
        }


        #region CONEXIÓN
        /// <summary>
        /// Regresa la coexion si es que se ha realizado
        /// </summary>
        /// <returns>regresa bool si esta o no conectado</returns>
        public bool GetConexion()
        {
            return Conexion;
        }
        /// <summary>
        /// Checa si la path seleccionada esta correcta conforme los archiovs k necesita admipaq
        /// </summary>
        /// <param name="PATH"></param>
        public void ConexionAdmipaq(string PATH)
        {
            List<string> ListArchivosAdmipaq = new List<string>();//para obtener los archivos admipaq
            ListArchivosAdmipaq = ModeloSDK.getListArchivosAdmi();//obtengo los archivos admipaq
            string[] files = Directory.GetFiles(@PATH, "*.dbf*");//obtengo todos los archivos de la carpeta seleccionada con extensión .dbf
            //debo de checar si existen o no los archivos
            int ArchivoValidado = 0;
            for (int i = 0; i < ListArchivosAdmipaq.Count; i++)//recorro la lista de los archivos necesarios
            {
                ArchivoValidado = 0;
                for (int j = 0; j < files.Length; j++)//recorro la lista delos archivo k estan en la carpeta
                {
                    string compara = PATH + "\\" + ListArchivosAdmipaq[i] + ".dbf";//genero la ruta con el nombre del archivo y su extension
                    if (compara.Equals(files[j]))
                    {
                        ArchivoValidado = 1;//si existe en la carpeta seleccionada solo me salgo del primero for y pongo mi bandera en 1
                        break;
                    }
                }
                if (ArchivoValidado == 0)//si mi bandera sale con cero del primer for significa que el archvo necesario no esta en la carpeta seleccionada
                { break; }
            }
            if (ArchivoValidado == 0)//si mi bandera sale con cero significa que no estan todos los archivos necesarios y pongo mi conexion como false y mando un mensaje
            {
                Conexion = false;
                MessageBox.Show("Favor de Verificar Su Conexión (Empresa no Válida)");
            }
            else//de lo contrartio pongo mi conexion como tru y mando un mensaje
            {
                Conexion = true;
                ModeloSDK.InicializaPath(PATH);
                MessageBox.Show("Conexión Exitosa", "Echo", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }
        #endregion




        #region CONSULTAS CRU
        /// <summary>
        /// consigue los documentos CRU
        /// </summary>
        /// <returns>regresa lo documentos leidos</returns>
        public List<Tipos_Datos_CRU.CRU> get_Documentos(string fechainicial, string fechafinal)
        {
            List<Tipos_Datos_CRU.get_agente> Listagentes = new List<Tipos_Datos_CRU.get_agente>();
            List<Tipos_Datos_CRU.Cliente_Proveedor> Listclietneproveeodres = new List<Tipos_Datos_CRU.Cliente_Proveedor>();
            List<Tipos_Datos_CRU.get_nom_concepto> Listnombconcepto = new List<Tipos_Datos_CRU.get_nom_concepto>();

            List<Tipos_Datos_CRU.CRU> ListDocmuentos = new List<Tipos_Datos_CRU.CRU>();//Creo el objeto donde regresare los documentos
            DataTable dtD = ModeloSDK.get_DocumentosCRU(fechainicial, fechafinal);//obtengo los docuemtos en un datatable
            if (dtD == null)//si tiene null siginifica que sucedio algun error 
                return ListDocmuentos;//regresa la lista vacia
            foreach (DataRow row in dtD.Rows)//recorro el databel
            {//si estan entre el filtro de fechas almacenalas en la lista si no no realizar anda
                //if (Convert.ToDateTime(row[7]) >= Convert.ToDateTime(fechainicial) && Convert.ToDateTime(row[7]) <= Convert.ToDateTime(fechafinal))
                //{
                Tipos_Datos_CRU.CRU newDocument = new Tipos_Datos_CRU.CRU()//crea el objeto para la lista
                {
                    IdDocumento = Convert.ToString(row[0]),
                    Serie = Convert.ToString(row[1]),
                    Folio = Convert.ToString(row[2]),
                    IDAgente = Convert.ToString(row[3]),
                    RazonSocial = Convert.ToString(row[4]),
                    FechaVencimiento = Convert.ToString(row[5]),
                    RFC = Convert.ToString(row[6]),
                    Fecha = Convert.ToString(row[7]),
                    Subtotal = (float)(double)row[8],
                    Total = (float)(double)row[9],
                    IVA = float.Parse(Math.Round((float)(double)row[9] - (float)(double)row[8], 2).ToString()),//redondea a 2 valores 
                    Pendiente = (float)(double)row[10],
                    TextoExtra1 = Convert.ToString(row[11]),
                    TextoExtra2 = Convert.ToString(row[12]),
                    TextoExtra3 = Convert.ToString(row[13]),
                    Cancelado = Convert.ToString(row[14]),
                    Impreso = Convert.ToString(row[15]),
                    Afectado = Convert.ToString(row[16]),
                    IDCliente = Convert.ToString(row[17]),
                    IDNombreConcepto = Convert.ToString(row[18]),
                   // NombreConcepto = ModeloSDK.GetNombreConcepto(Convert.ToString(row[18])),
                    TotalUnidades = (float)(double)row[19],
                    CIDDOCUM02 = Convert.ToString(row[20]).Trim(),
                    CIDCONCE01 = Convert.ToString(row[21]).Trim()

                };
                int posicion = -5;
                for (int i = 0; i < Listnombconcepto.Count; i++)
                {
                    if (Listnombconcepto[i].IDNombreConcepto == newDocument.IDNombreConcepto)
                    {
                        posicion = i;
                        break;
                    }
                }
                if (posicion < 0)
                {
                    Tipos_Datos_CRU.get_nom_concepto new_conept = new Tipos_Datos_CRU.get_nom_concepto
                    {
                        IDNombreConcepto = newDocument.IDNombreConcepto,
                        nombre_concepto = ModeloSDK.GetNombreConcepto(Convert.ToString(newDocument.IDNombreConcepto))
                    };
                    Listnombconcepto.Add(new_conept);
                }
                else
                {
                    newDocument.NombreConcepto = Listnombconcepto[posicion].nombre_concepto;
                }
                //List<Tipos_Datos_CRU.get_agente> Listagentes = new List<Tipos_Datos_CRU.get_agente>();
                posicion = -5;
                Tipos_Datos_CRU.get_agente get_agente = new Tipos_Datos_CRU.get_agente();
                for (int i = 0; i < Listagentes.Count; i++)//checa si ya tengo ese dato si lo tengo no lo mandes a pedir y solo
                {
                    if (Listagentes[i].IDAgente == newDocument.IDAgente)
                    {
                        posicion = i;
                        break;
                    }
                }
                if (posicion < 0)
                {
                    DataRow rowagente = ModeloSDK.GETNombreAgente(newDocument.IDAgente);//obtengo el nombre y coidgo del agente
                    if (rowagente != null)//si regresa null significa que existio algun error y por lo cual no ara nada
                    {
                        Tipos_Datos_CRU.get_agente newagente = new Tipos_Datos_CRU.get_agente
                        {
                            CodigoAgente = Convert.ToString(rowagente[0]),
                            NombreAgente = Convert.ToString(rowagente[1]),
                            IDAgente = newDocument.IDAgente.Trim()
                        };
                        Listagentes.Add(newagente);

                        newDocument.CodigoAgente = Convert.ToString(rowagente[0]);
                        newDocument.NombreAgente = Convert.ToString(rowagente[1]);
                    }
                }
                else
                {
                    newDocument.CodigoAgente = Listagentes[posicion].CodigoAgente;
                    newDocument.NombreAgente = Listagentes[posicion].NombreAgente;
                }
                //se necesitan sacar los datos del cliente/proveedor
                posicion = -5;

                for (int i = 0; i < Listclietneproveeodres.Count; i++)
                {
                    if (Listclietneproveeodres[i].IDCliente == newDocument.IDCliente)
                    {
                        posicion = i;
                        break;
                    }
                }

                if (posicion < 0)
                {
                    DataRow rowClientePRoveedor = ModeloSDK.GETCLientePRoveedor(newDocument.IDCliente);
                    if (rowClientePRoveedor != null)//Saca ñps datos del cliente/proveedor si es que existe
                    {
                        Tipos_Datos_CRU.Cliente_Proveedor Proveedor = new Tipos_Datos_CRU.Cliente_Proveedor()
                        {
                            CodigoCliente = Convert.ToString(rowClientePRoveedor[0]),
                            RazonSocial = Convert.ToString(rowClientePRoveedor[1]),
                            ValorClasificación1 = Convert.ToString(rowClientePRoveedor[2]),
                            ValorClasificación2 = Convert.ToString(rowClientePRoveedor[3]),
                            ValorClasificación3 = Convert.ToString(rowClientePRoveedor[4]),
                            Clasificación1 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowClientePRoveedor[2])),
                            Clasificación2 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowClientePRoveedor[3])),
                            Clasificación3 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowClientePRoveedor[4])),
                            IDCliente = newDocument.IDCliente

                        };
                        Listclietneproveeodres.Add(Proveedor);
                        newDocument.proveedor = Proveedor;
                    }
                    else
                    {
                        newDocument.proveedor = null;//sino existen almacena un null
                    }
                }
                else
                {
                    newDocument.proveedor = Listclietneproveeodres[posicion];
                }
                /*
                   //comienza a sacar los datos de movimientos
                    DataTable dtM = ModeloSDK.get_MovimientosCRU(newDocument.IdDocumento);
                    List<Tipos_Datos_CRU.Movimientos> ListMovimientos = new List<Tipos_Datos_CRU.Movimientos>();
                    foreach (DataRow rowM in dtM.Rows)
                    {
                        Tipos_Datos_CRU.Movimientos newMovimiento = new Tipos_Datos_CRU.Movimientos()
                        {
                            IDProducto=Convert.ToString(rowM[0]),
                            CantidadProducto = Convert.ToString(rowM[1]),
                            PrecioProducto = Convert.ToString(rowM[2]),
                            Importe = (float)(double)rowM[3],
                            Total = (float)(double)rowM[4],
                            IVA = float.Parse(Math.Round((float)(double)rowM[4] - (float)(double)rowM[3], 2).ToString())
                        };
                        //sacar los productos
                        DataRow rowProducto = ModeloSDK.getProductos(newMovimiento.IDProducto);
                        if (rowProducto != null)
                        {
                            Tipos_Datos_CRU.Producto newProducto = new Tipos_Datos_CRU.Producto()
                            {
                                codigo = Convert.ToString(rowProducto[0]),
                                Descripcion = Convert.ToString(rowProducto[1]),
                                ValorClasificación1 = Convert.ToString(rowProducto[2]),
                                ValorClasificación2 = Convert.ToString(rowProducto[3]),
                                ValorClasificación3 = Convert.ToString(rowProducto[4]),
                                Clasifiacion1 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowProducto[2])),
                                Clasificacion2 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowProducto[3])),
                                Clasificacion3 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowProducto[4]))
                            };
                            newMovimiento.producto = newProducto;
                        }
                        else newMovimiento.producto = null;//como tiene producto o marco algun error pon el producto como null
                        ListMovimientos.Add(newMovimiento);
                    }

                    newDocument.Listmovimiento = ListMovimientos;//ya se agregaron la lista de los movimientos
                    //temina y guarda los datos de los movimientos
                   */

                
                ListDocmuentos.Add(newDocument);
                //}//else MessageBox.Show("fecha");
            }
            return ListDocmuentos;//regresa la lista
        }



        /// <summary>
        /// como se obtienen todo los datos de golpe ahora divide todos en su lugar correspondiente
        /// </summary>
        /// <param name="ListDocumentos">lista de todos los datos obenidos</param>
        /// <param name="RFCPublico">filtor para RFC publico</param>
        /// <param name="RFCOl">rfc ol</param>
        /// <param name="RFCAnji">rfca para anji</param>
        /// <returns>regresa un objeto con los datos dividiso</returns>
        public Tipos_Datos_CRU.ListDatosCRU filtro_indicadores_tipo(List<Tipos_Datos_CRU.CRU> ListDocumentos,string RFCPublico,string RFCOl,string RFCAnji)
        {
            Tipos_Datos_CRU.ListDatosCRU filtro_suc = new Tipos_Datos_CRU.ListDatosCRU();

            List<Tipos_Datos_CRU.CRU> facturas = new List<Tipos_Datos_CRU.CRU>();           //ok
             List<Tipos_Datos_CRU.CRU> facturas_rfc_publico = new List<Tipos_Datos_CRU.CRU>();//ok
             List<Tipos_Datos_CRU.CRU> facturas_rfc_ol = new List<Tipos_Datos_CRU.CRU>();    //ok
             List<Tipos_Datos_CRU.CRU> compras = new List<Tipos_Datos_CRU.CRU>();            //ok
             List<Tipos_Datos_CRU.CRU> compras_rfc_anji = new List<Tipos_Datos_CRU.CRU>(); //ok
             List<Tipos_Datos_CRU.CRU> abonos = new List<Tipos_Datos_CRU.CRU>();         //ok
             List<Tipos_Datos_CRU.CRU> abonos_rfc_publico = new List<Tipos_Datos_CRU.CRU>();//ok
             List<Tipos_Datos_CRU.CRU> abonos_ol = new List<Tipos_Datos_CRU.CRU>();//ok
             List<Tipos_Datos_CRU.CRU> pagos_proveedor = new List<Tipos_Datos_CRU.CRU>();//ok
             List<Tipos_Datos_CRU.CRU> pagos_proveedor_rfc_anji = new List<Tipos_Datos_CRU.CRU>();//ok
             List<Tipos_Datos_CRU.CRU> prestamos= new List<Tipos_Datos_CRU.CRU>();
             List<Tipos_Datos_CRU.CRU> ingreso_traspaso = new List<Tipos_Datos_CRU.CRU>();
             List<Tipos_Datos_CRU.CRU> ingreso_dev_garantia = new List<Tipos_Datos_CRU.CRU>();

             List<Tipos_Datos_CRU.CRU> abonos_zona_norte= new List<Tipos_Datos_CRU.CRU>();
             List<Tipos_Datos_CRU.CRU> abonos_zona_centro= new List<Tipos_Datos_CRU.CRU>();
             List<Tipos_Datos_CRU.CRU> abonos_zona_sur = new List<Tipos_Datos_CRU.CRU>();


             filtro_suc.facturas = facturas;
             filtro_suc.facturas_rfc_publico = facturas_rfc_publico;
             filtro_suc.facturas_rfc_ol = facturas_rfc_ol;
             filtro_suc.compras = compras;
             filtro_suc.compras_rfc_anji = compras_rfc_anji;
             filtro_suc.abonos = abonos;
             filtro_suc.abonos_rfc_publico = abonos_rfc_publico;
             filtro_suc.abonos_ol = abonos_ol;
             filtro_suc.pagos_proveedor = pagos_proveedor;
             filtro_suc.pagos_proveedor_rfc_anji = pagos_proveedor_rfc_anji;
             filtro_suc.prestamos = prestamos;
             filtro_suc.ingreso_traspaso = ingreso_traspaso;
             filtro_suc.ingreso_dev_garantia = ingreso_dev_garantia;

             filtro_suc.abonos_zona_norte = abonos_zona_norte;
             filtro_suc.abonos_zona_centro = abonos_zona_centro;
             filtro_suc.abonos_zona_sur = abonos_zona_sur;
            for (int i = 0; i < ListDocumentos.Count; i++)
            {
                Tipos_Datos_CRU.CRU new_data= new Tipos_Datos_CRU.CRU();
                new_data = ListDocumentos[i];
                if (ListDocumentos[i].CIDDOCUM02.Trim() == "4" && ListDocumentos[i].CIDCONCE01.Trim() == "3007")
                {//significa que es una factura 
                    facturas.Add(new_data);//compara pa meter los filtro por ol y por publico
                    if (ListDocumentos[i].RFC.ToUpper().Trim().Equals(RFCOl.ToUpper().Trim()))
                    {
                        facturas_rfc_ol.Add(new_data);
                    }
                    else { facturas_rfc_publico.Add(new_data); }
                }//si no es factura entonces checa si es compra
                else if (ListDocumentos[i].CIDDOCUM02.Trim() == "19" && ListDocumentos[i].CIDCONCE01.Trim() == "21")
                {
                    compras.Add(new_data);//compara por el filtro de anji para el filtro de compras
                    if (ListDocumentos[i].RFC.ToUpper().Trim().Equals(RFCAnji.ToUpper().Trim()))
                    {
                        compras_rfc_anji.Add(new_data);
                    }
                }//si no es compra entonces checa si es un abono
                else if (ListDocumentos[i].CIDDOCUM02.Trim() == "12" && ListDocumentos[i].CIDCONCE01.Trim() == "13")
                {
                    abonos.Add(new_data);
                    if (ListDocumentos[i].RFC.ToUpper().Trim().Equals(RFCOl.ToUpper().Trim()))
                    {
                        abonos_ol.Add(new_data);
                    }
                    else { abonos_rfc_publico.Add(new_data); }
                    //checa de que zona es cada uno
                    if (ListDocumentos[i].proveedor.Clasificación1.Trim().ToUpper().Equals("ZONA NORTE".ToUpper()))
                    {
                        abonos_zona_norte.Add(ListDocumentos[i]);
                    }
                    else if (ListDocumentos[i].proveedor.Clasificación1.Trim().ToUpper().Equals("ZONA CENTRO".ToUpper()))
                    {
                        abonos_zona_centro.Add(ListDocumentos[i]);
                    }
                    else
                    {
                        abonos_zona_sur.Add(ListDocumentos[i]);
                    }





                }//si no es un abono entonces checa si es un pago proveedor
                else if (ListDocumentos[i].CIDDOCUM02.Trim() == "23" && ListDocumentos[i].CIDCONCE01.Trim() == "25")
                {
                    pagos_proveedor.Add(new_data);
                    if (ListDocumentos[i].RFC.ToUpper().Trim().Equals(RFCAnji.ToUpper().Trim()))
                    {
                        pagos_proveedor_rfc_anji.Add(new_data);
                    }
                }//si no es pago proveedor entonces  chec prestamos
                else if (ListDocumentos[i].CIDDOCUM02.Trim() == "12" && ListDocumentos[i].CIDCONCE01.Trim() == "3011")
                {
                    prestamos.Add(new_data);
                }//si no es prestamo entonces checa ingreso traspaso
                else if (ListDocumentos[i].CIDDOCUM02.Trim() == "12" && ListDocumentos[i].CIDCONCE01.Trim() == "3010")
                {
                    ingreso_traspaso.Add(new_data);
                }//si no es ingreso traspaso entonces checa ingreso devolucion garantia
                else if (ListDocumentos[i].CIDDOCUM02.Trim() == "12" && ListDocumentos[i].CIDCONCE01.Trim() == "3012")
                {
                    ingreso_dev_garantia.Add(new_data);
                }
               // else { MessageBox.Show("no"); }


            }



            return filtro_suc;            
        }


      
        

        
        #endregion


     

        


        #region CONSULTAS productos
        /// <summary>
        /// consigue los documentos CXP
        /// </summary>
        /// <returns>regresa los documentos leidos</returns>
        //public List<Tipos_Datos_CRU.Producto> get_Productos(string fechainicial, string fechafinal)
        //{
        //    List<Tipos_Datos_CRU.Producto> ListDocmuentos = new List<Tipos_Datos_CRU.Producto>();//Creo el objeto donde regresare los documentos
        //    DataTable dtD = ModeloSDK.getProductos_fecha(fechainicial, fechafinal);//obtengo los docuemtos en un datatable
        //    if (dtD == null)//si tiene null siginifica que sucedio algun error 
        //        return ListDocmuentos;//regresa la lista vacia
        //    foreach (DataRow row in dtD.Rows)//recorro el datatable
        //    {
        //        si estan entre el filtro de fechas almacenalas en la lista si no no realizar anda

        //        Tipos_Datos_CRU.Producto newDocument = new Tipos_Datos_CRU.Producto()//crea el objeto para la lista
        //        {
        //            codigo = Convert.ToString(row[0]),
        //            Descripcion = Convert.ToString(row[1]),
        //            ValorClasificación1 = Convert.ToString(row[2]),
        //            ValorClasificación2 = Convert.ToString(row[3]),                   
        //            ValorClasificación3 = ""                  
        //        };

        //        newDocument.Clasifiacion1 = ModeloSDK.getProductosclasificacion_1(newDocument.ValorClasificación1);
        //        string resp= Convert.ToString(ModeloSDK.getclasificacion(newDocument.ValorClasificación2));
        //        string[] words = resp.Split(',');
        //        newDocument.Clasificacion2 = words[0];
        //        newDocument.Clasificacion3 = words[1];

        //        ListDocmuentos.Add(newDocument);
        //    }
        //    return ListDocmuentos;//regresa la lista
        //}
        #endregion



       

       



        
    }
}
