using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IndicadoresISEL.Modelo;

namespace IndicadoresISEL.Controlador
{
    class Controlador_Impresion
    {
        Modelo_Impresion modeloimpresion;//objeto para comunicarse con el modelo de impresion
        /// <summary>
        /// constructor para controlador de impresion
        /// </summary>
        public Controlador_Impresion()
        {
            modeloimpresion = new Modelo_Impresion();
        }


        public void ImpresionCRUFacturas(List<Tipos_Datos_CRU.FacturasCRU> ListFactrurasCRU, string fechas, string path, List<Tipos_Datos_CRU.FacturasCRU> ListFactrurasCRUFiltroRFCPublico, List<Tipos_Datos_CRU.FacturasCRU> ListFactrurasCRUFiltroRFCOL)
        {
            modeloimpresion.ImpresionCRUFacturas(ListFactrurasCRU, fechas, path, ListFactrurasCRUFiltroRFCPublico, ListFactrurasCRUFiltroRFCOL);

        }

        public void impresion_productos(List<Tipos_Datos_CRU.Producto> listaproductos, string fecha_ini, string fecha_fin)
        {
            modeloimpresion.ImpresionTablaProductos(listaproductos, fecha_ini, fecha_fin);
        }

        public void impresion_movimientos_productos(List<Tipos_Datos_CRU.Movimientos_Cuentas> lista, string fechas, string fecha_titulo, string path)
        {
            modeloimpresion.Reporte_Compras(lista, fechas, fecha_titulo, path);
        }

        public void excel_import(List<Tipos_Datos_CRU.FacturasCRU> ListDocmuentos, List<Tipos_Datos_CRU.FacturasCRU> list_rfc_publico, List<Tipos_Datos_CRU.FacturasCRU> list_rfc_ol)
        {
            modeloimpresion.excel_import(ListDocmuentos, list_rfc_publico, list_rfc_ol);
        
        }


    }
}
