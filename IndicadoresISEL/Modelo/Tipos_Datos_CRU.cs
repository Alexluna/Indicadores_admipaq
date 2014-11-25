using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IndicadoresISEL.Modelo
{
    class Tipos_Datos_CRU
    {
        #region FACTURAS CRU
        public class get_nom_concepto {
            public string IDNombreConcepto { get; set; }
            public string nombre_concepto { get; set; }
        }

        public class get_agente {
            public string IDAgente { get; set; }
            public string NombreAgente { get; set; }
            public string CodigoAgente { get; set; }
        }

        

        public class ListDatosCRU
        {
            public List<CRU> facturas { get; set; }
            public List<CRU> facturas_rfc_publico { get; set; }
            public List<CRU> facturas_rfc_ol { get; set; }
            public List<CRU> compras { get; set; }
            public List<CRU> compras_rfc_anji { get; set; }
            public List<CRU> abonos { get; set; }
            public List<CRU> abonos_rfc_publico { get; set; }
            public List<CRU> abonos_ol { get; set; }
            public List<CRU> abonos_zona_norte { get; set; }
            public List<CRU> abonos_zona_centro { get; set; }
            public List<CRU> abonos_zona_sur { get; set; }
            public List<CRU> pagos_proveedor { get; set; }
            public List<CRU> pagos_proveedor_rfc_anji { get; set; }
            public List<CRU> prestamos { get; set; }
            public List<CRU> ingreso_traspaso { get; set; }
            public List<CRU> ingreso_dev_garantia { get; set; }
        }

        public class CRU
        {
            public string CIDDOCUM02 { get; set; }
            public string CIDCONCE01 { get; set; }
            public string IdDocumento { get; set; }
            public string Fecha { get; set; }
            public string Serie { get; set; }
            public string Folio { get; set; }
            public string NombreAgente { get; set; }
            public string CodigoAgente { get; set; }
            public string IDAgente { get; set; }
            public string RazonSocial { get; set; }
            public string FechaVencimiento { get; set; }
            public string RFC { get; set; }
            public float Subtotal { get; set; }
            public float IVA { get; set; }
            public float Total { get; set; }
            public float Pendiente { get; set; }
            public float TotalUnidades { get; set; }
            public string TextoExtra1 { get; set; }
            public string TextoExtra2 { get; set; }
            public string TextoExtra3 { get; set; }
            public string Afectado { get; set; }
            public string Impreso { get; set; }
            public string Cancelado { get; set; }
            public string Clasificacion1 { get; set; }
            public string Clasificacion2 { get; set; }
            public string Clasificacion3 { get; set; }
            public string NombreConcepto { get; set; }
            public string CodigoCliente { get; set; }
            public string IDCliente { get; set; }
            public string IDNombreConcepto { get; set; }
            public string Referencia { get; set; }
            public Cliente_Proveedor proveedor { get; set; }
            public List<Movimientos> Listmovimiento { get; set; }

        }
        #endregion

        


        public class Movimientos_Cuentas// CUENTAS POR PAGAR 
        {

            public string fecha { get; set; }
            public string folio { get; set; }
            public string Proveedor { get; set; }
            public string Proveedor_codigo { get; set; }
            public string IDProducto { get; set; }
            // public string NombreProducto { get; set; } //ver si es necesario si llegara hacer necesario se tiene que hacer una busqueda de producto y sacar solo el nombre
            public string PrecioProducto { get; set; }//precio unitario
            public string CantidadProducto { get; set; }//cuantos productos
            public float Total { get; set; }
            public float Subtotal { get; set; }
            public float Importe { get; set; }
            public float IVA { get; set; }
            public string ID_doc { get; set; }
            public string producto_codigo { get; set; }
            public string producto_nombre { get; set; }
            public int semana { get; set; }
            public string Clasificacion_1_producto { get; set; }
            public string Valor_Clasificacion_1_producto { get; set; }
            public string Clasificacion_2_producto { get; set; }
            public string Valor_Clasificacion_2_producto { get; set; }
            public string Clasificacion_1_proveedor { get; set; }
            public string Valor_Clasificacion_1_proveedor { get; set; }
            public string Clasificacion_2_proveedor { get; set; }
            public string Valor_Clasificacion_2_proveedor { get; set; }
            public string pendiente { get; set; }
            public string IDCliente { get; set; }
            public Cliente_Proveedor proveedor { get; set; }
            public List<Movimientos> Listmovimiento { get; set; }


        }


        public class Movimientos
        {

            public string fecha { get; set; }
            public string IDProducto { get; set; }
            public string NombreProducto { get; set; } //ver si es necesario si llegara hacer necesario se tiene que hacer una busqueda de producto y sacar solo el nombre
            public string PrecioProducto { get; set; }//precio unitario
            public string CantidadProducto { get; set; }//cuantos productos
            public float Total { get; set; }
            public float Subtotal { get; set; }
            public float Importe { get; set; }
            public float IVA { get; set; }
            public string ID_doc { get; set; }


            public Producto producto { get; set; }


        }


        public class Producto
        {
            public string codigo { get; set; }
            public string Descripcion { get; set; }
            public string Clasifiacion1 { get; set; }
            public string ValorClasificación1 { get; set; }
            public string Clasificacion2 { get; set; }
            public string ValorClasificación2 { get; set; }
            public string Clasificacion3 { get; set; }
            public string ValorClasificación3 { get; set; }
        }

        public class Cliente_Proveedor
        {
            public string IDCliente { get; set; }
            public string CodigoCliente { get; set; }
            public string RazonSocial { get; set; }
            public string Clasificación1 { get; set; }
            public string ValorClasificación1 { get; set; }
            public string Clasificación2 { get; set; }
            public string ValorClasificación2 { get; set; }
            public string Clasificación3 { get; set; }
            public string ValorClasificación3 { get; set; }
        }


        
    }
}
