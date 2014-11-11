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
        public List<Tipos_Datos_CRU.FacturasCRU> get_Documentos(string fechainicial, string fechafinal)
        {
            
            List<Tipos_Datos_CRU.FacturasCRU> ListDocmuentos = new List<Tipos_Datos_CRU.FacturasCRU>();//Creo el objeto donde regresare los documentos
            DataTable dtD = ModeloSDK.get_DocumentosCRU(fechainicial, fechafinal);//obtengo los docuemtos en un datatable
            if (dtD == null)//si tiene null siginifica que sucedio algun error 
                return ListDocmuentos;//regresa la lista vacia
            foreach (DataRow row in dtD.Rows)//recorro el databel
            {//si estan entre el filtro de fechas almacenalas en la lista si no no realizar anda
                //if (Convert.ToDateTime(row[7]) >= Convert.ToDateTime(fechainicial) && Convert.ToDateTime(row[7]) <= Convert.ToDateTime(fechafinal))
                //{
                Tipos_Datos_CRU.FacturasCRU newDocument = new Tipos_Datos_CRU.FacturasCRU()//crea el objeto para la lista
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
                    NombreConcepto = ModeloSDK.GetNombreConcepto(Convert.ToString(row[18])),
                    TotalUnidades = (float)(double)row[19]

                };

                DataRow rowagente = ModeloSDK.GETNombreAgente(newDocument.IDAgente);//obtengo el nombre y coidgo del agente
                if (rowagente != null)//si regresa null significa que existio algun error y por lo cual no ara nada
                {
                    newDocument.CodigoAgente = Convert.ToString(rowagente[0]);
                    newDocument.NombreAgente = Convert.ToString(rowagente[1]);
                }
                //se necesitan sacar los datos del cliente/proveedor

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
                        Clasificación3 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowClientePRoveedor[4]))

                    };
                    newDocument.proveedor = Proveedor;
                }
                else
                {
                    newDocument.proveedor = null;//sino existen almacena un null
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



        public List<Tipos_Datos_CRU.FacturasCRU> FiltroRFCCRU(List<Tipos_Datos_CRU.FacturasCRU> ListDocumentos, string RFCFiltro)
        {
            List<Tipos_Datos_CRU.FacturasCRU> ListFiltroRFC = new List<Tipos_Datos_CRU.FacturasCRU>();
            for (int i = 0; i < ListDocumentos.Count; i++)
            {
                if (ListDocumentos[i].RFC.ToUpper().Trim().Equals(RFCFiltro.ToUpper().Trim()))
                {
                    ListFiltroRFC.Add(ListDocumentos[i]);
                }
            }
            return ListFiltroRFC;
        }

        #endregion


        #region CONSULTAS CRU ABONOS
        /// <summary>
        /// consigue los documentos CRU
        /// </summary>
        /// <returns>regresa lo documentos leidos</returns>
        public List<Tipos_Datos_CRU.FacturasCRU> get_AbonosCRU(string fechainicial, string fechafinal)
        {
            List<Tipos_Datos_CRU.FacturasCRU> ListDocmuentos = new List<Tipos_Datos_CRU.FacturasCRU>();//Creo el objeto donde regresare los documentos
            DataTable dtD = ModeloSDK.get_AbonosCRU(fechainicial, fechafinal);//obtengo los docuemtos en un datatable
            if (dtD == null)//si tiene null siginifica que sucedio algun error 
                return ListDocmuentos;//regresa la lista vacia
            foreach (DataRow row in dtD.Rows)//recorro el databel
            {//si estan entre el filtro de fechas almacenalas en la lista si no no realizar anda
                //if (Convert.ToDateTime(row[7]) >= Convert.ToDateTime(fechainicial) && Convert.ToDateTime(row[7]) <= Convert.ToDateTime(fechafinal))
                //{
                Tipos_Datos_CRU.FacturasCRU newDocument = new Tipos_Datos_CRU.FacturasCRU()//crea el objeto para la lista
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
                    NombreConcepto = ModeloSDK.GetNombreConcepto(Convert.ToString(row[18])),
                    TotalUnidades = (float)(double)row[19]

                };

                DataRow rowagente = ModeloSDK.GETNombreAgente(newDocument.IDAgente);//obtengo el nombre y coidgo del agente
                if (rowagente != null)//si regresa null significa que existio algun error y por lo cual no ara nada
                {
                    newDocument.CodigoAgente = Convert.ToString(rowagente[0]);
                    newDocument.NombreAgente = Convert.ToString(rowagente[1]);
                }
                //se necesitan sacar los datos del cliente/proveedor

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
                        Clasificación3 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowClientePRoveedor[4]))

                    };
                    newDocument.proveedor = Proveedor;
                }
                else
                {
                    newDocument.proveedor = null;//sino existen almacena un null
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



        #region CONSULTAS MOVIMIENTOS

        /// <summary>
        /// consigue los movimientos de compras por fecha
        /// </summary>
        /// <returns>regresa lo documentos leidos</returns>
        public List<Tipos_Datos_CRU.Movimientos_Cuentas> get_Movimientos_Compras(string fechainicial, string fechafinal)
        {
            List<Tipos_Datos_CRU.Movimientos_Cuentas> ListMovimientos_Compras = new List<Tipos_Datos_CRU.Movimientos_Cuentas>();//Creo el objeto donde regresare los movimientos de las compras 
            DataTable dtD = ModeloSDK.get_documentos_Compras(fechainicial, fechafinal);//obtengo los movimientos en un datatable

            if (dtD == null)//si tiene null siginifica que sucedio algun error 
                return ListMovimientos_Compras;//regresa la lista vacia

            foreach (DataRow row in dtD.Rows)//recorro el databel
            {
                string[] fecha = Convert.ToString(row[0]).Split(' ');//obtengo la fecha de la lista 
                //  string[] fecha_partes = fecha[0].Split('/');//separa la fecha por dia mes y año [DIA][MES][AÑO]
                Tipos_Datos_CRU.Movimientos_Cuentas newDocument = new Tipos_Datos_CRU.Movimientos_Cuentas()//crea el objeto para la lista
                {
                    fecha = fecha[0].TrimEnd(' '),
                    Proveedor = Convert.ToString(row[1]).TrimEnd(' '),
                    CantidadProducto = Convert.ToString(row[2]).TrimEnd(' '),
                    PrecioProducto = Convert.ToString(row[3]).TrimEnd(' '),
                    Subtotal = (float)(double)row[4],
                    IVA = (float)(double)row[5],
                    Total = (float)(double)row[6],
                    ID_doc = Convert.ToString(row[7]).TrimEnd(' '),
                    IDCliente = Convert.ToString(row[8]).TrimEnd(' '),
                    pendiente = Convert.ToString(row[9]).TrimEnd(' '),
                    folio = Convert.ToString(row[10]).TrimEnd(' ')
                };


                DataRow rowproducto = ModeloSDK.get_Movimientos_Productos(newDocument.ID_doc);

                if (rowproducto != null)//si regresa null significa que existio algun error y por lo cual no ara nada
                {
                    newDocument.producto_codigo = Convert.ToString(rowproducto[0]).TrimEnd(' ');
                    newDocument.producto_nombre = Convert.ToString(rowproducto[1]).TrimEnd(' ');
                }



                string resp = Convert.ToString(ModeloSDK.getclasificacion(Convert.ToString(rowproducto[2])));
                string[] words = resp.Split(',');
                newDocument.Valor_Clasificacion_1_producto = words[0].TrimEnd(' ');
                newDocument.Clasificacion_1_producto = words[1].TrimEnd(' ');


                string resp2 = Convert.ToString(ModeloSDK.getclasificacion(Convert.ToString(rowproducto[3])));
                string[] words2 = resp.Split(',');
                newDocument.Valor_Clasificacion_2_producto = words2[0].TrimEnd(' ');
                newDocument.Clasificacion_2_producto = words2[1].TrimEnd(' ');

                DataRow rowprovedor = ModeloSDK.get_codigo_proveedor(newDocument.IDCliente);//obtengo el codigo del provedor  sus valores 

                if (rowprovedor != null)
                {
                    newDocument.Proveedor_codigo = Convert.ToString(rowprovedor[0]);

                    string response = Convert.ToString(ModeloSDK.getclasificacion(Convert.ToString(rowprovedor[1])));
                    string[] words_ = response.Split(',');
                    newDocument.Valor_Clasificacion_1_proveedor = words_[0].TrimEnd(' ');
                    newDocument.Clasificacion_1_proveedor = words_[1].TrimEnd(' ');


                    string responce2 = Convert.ToString(ModeloSDK.getclasificacion(Convert.ToString(rowprovedor[2])));
                    string[] words2_ = responce2.Split(',');
                    newDocument.Valor_Clasificacion_2_proveedor = words2_[0].TrimEnd(' ');
                    newDocument.Clasificacion_2_proveedor = words2_[1].TrimEnd(' ');

                }


                //comienza a sacar los datos de movimientos
                DataTable dtM = ModeloSDK.get_MovimientosCRU(newDocument.ID_doc);
                List<Tipos_Datos_CRU.Movimientos> ListMovimientos = new List<Tipos_Datos_CRU.Movimientos>();
                foreach (DataRow rowM in dtM.Rows)
                {
                    Tipos_Datos_CRU.Movimientos newMovimiento = new Tipos_Datos_CRU.Movimientos()
                    {
                        IDProducto = Convert.ToString(rowM[0]).TrimEnd(' '),
                        CantidadProducto = Convert.ToString(rowM[1]).TrimEnd(' '),
                        PrecioProducto = Convert.ToString(rowM[2]).TrimEnd(' '),
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
                            codigo = Convert.ToString(rowProducto[0]).TrimEnd(' '),
                            Descripcion = Convert.ToString(rowProducto[1]).TrimEnd(' '),
                            ValorClasificación1 = Convert.ToString(rowProducto[2]).TrimEnd(' '),
                            ValorClasificación2 = Convert.ToString(rowProducto[3]).TrimEnd(' '),
                            ValorClasificación3 = Convert.ToString(rowProducto[4]).TrimEnd(' '),
                            Clasifiacion1 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowProducto[2]).TrimEnd(' ')),
                            Clasificacion2 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowProducto[3]).TrimEnd(' ')),
                            Clasificacion3 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowProducto[4]).TrimEnd(' '))
                        };
                        newMovimiento.producto = newProducto;
                    }
                    else newMovimiento.producto = null;//como tiene producto o marco algun error pon el producto como null
                    ListMovimientos.Add(newMovimiento);
                }

                newDocument.Listmovimiento = ListMovimientos;//ya se agregaron la lista de los movimientos
                //temina y guarda los datos de los movimientos

                ListMovimientos_Compras.Add(newDocument);
            }
            return ListMovimientos_Compras;//regresa la lista
        }
        #endregion



        #region CONSULTAS CRU CXC
        /// <summary>
        /// consigue los documentos CRU
        /// </summary>
        /// <returns>regresa lo documentos leidos</returns>
        public List<Tipos_Datos_CRU.Abonos_cxc> get_Documentos_CXC(string fechainicial, string fechafinal)
        {
            List<Tipos_Datos_CRU.Abonos_cxc> ListDocmuentos = new List<Tipos_Datos_CRU.Abonos_cxc>();//Creo el objeto donde regresare los documentos
            DataTable dtD = ModeloSDK.get_DocumentosCRU_cxc(fechainicial, fechafinal);//obtengo los docuemtos en un datatable
            if (dtD == null)//si tiene null siginifica que sucedio algun error 
                return ListDocmuentos;//regresa la lista vacia
            foreach (DataRow row in dtD.Rows)//recorro el databel
            {//si estan entre el filtro de fechas almacenalas en la lista si no no realizar anda
                //if (Convert.ToDateTime(row[7]) >= Convert.ToDateTime(fechainicial) && Convert.ToDateTime(row[7]) <= Convert.ToDateTime(fechafinal))
                //{
                Tipos_Datos_CRU.Abonos_cxc newDocument = new Tipos_Datos_CRU.Abonos_cxc()//crea el objeto para la lista
                {
                    Fecha = Convert.ToString(row[0]),
                    Folio = Convert.ToString(row[1]),
                    IDCliente = Convert.ToString(row[2]),
                    RazonSocial = Convert.ToString(row[3]),
                    RFC = Convert.ToString(row[4]),
                    Total = (float)(double)row[5],
                    Pendiente = (float)(double)row[6],
                    Referencia = Convert.ToString(row[7]),
                    Cuenta = Convert.ToString(row[8]),
                    IDAgente = Convert.ToString(row[9]),

                };

                DataRow rowagente = ModeloSDK.GETNombreAgente(newDocument.IDAgente);//obtengo el nombre y coidgo del agente
                if (rowagente != null)//si regresa null significa que existio algun error y por lo cual no ara nada
                {
                    newDocument.CodigoAgente = Convert.ToString(rowagente[0]);
                    newDocument.NombreAgente = Convert.ToString(rowagente[1]);
                }
                //se necesitan sacar los datos del cliente/proveedor

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
                        Clasificación3 = ModeloSDK.GetValoresClasificacionClientesPRoveedores(Convert.ToString(rowClientePRoveedor[4]))

                    };
                    newDocument.proveedor = Proveedor;
                }
                else
                {
                    newDocument.proveedor = null;//sino existen almacena un null
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



        public List<Tipos_Datos_CRU.Abonos_cxc> FiltroRFCCRU(List<Tipos_Datos_CRU.Abonos_cxc> ListDocumentos, string RFCFiltro)
        {
            List<Tipos_Datos_CRU.Abonos_cxc> ListFiltroRFC = new List<Tipos_Datos_CRU.Abonos_cxc>();
            for (int i = 0; i < ListDocumentos.Count; i++)
            {
                if (ListDocumentos[i].RFC.ToUpper().Equals(RFCFiltro.ToUpper()))
                {
                    ListFiltroRFC.Add(ListDocumentos[i]);
                }
            }
            return ListFiltroRFC;
        }

        #endregion
    }
}
