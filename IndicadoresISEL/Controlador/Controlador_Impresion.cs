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







        /// <summary>
        /// impresion excel para los indicadores de CRU
        /// </summary>
        /// <param name="ListDocmuentos"></param>
        public void excel_importCRU(Tipos_Datos_CRU.ListDatosCRU ListDocmuentos)
        {
            modeloimpresion.excel_importCRUs(ListDocmuentos);

        }

        /// <summary>
        /// impresion excel para los indicadores de CRU
        /// </summary>
        /// <param name="ListDocmuentos"></param>
        public void excel_importOL(Tipos_Datos_CRU.ListDatosOL ListDocmuentos)
        {
            modeloimpresion.excel_importOL(ListDocmuentos);

        }


        /// <summary>
        /// impresion excel para los indicadores de CRU
        /// </summary>
        /// <param name="ListDocmuentos"></param>
        public void excel_importISEL(Tipos_Datos_CRU.ListDatosISEL ListDocmuentos)
        {
            modeloimpresion.excel_importISEL(ListDocmuentos);

        }

        /// <summary>
        /// impresion excel para los indicadores de CRU
        /// </summary>
        /// <param name="ListDocmuentos"></param>
        public void excel_importMANUEL(Tipos_Datos_CRU.ListDatosMANUEL ListDocmuentos)
        {
            modeloimpresion.excel_importMANUEL(ListDocmuentos);

        }


    }
}
