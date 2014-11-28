using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IndicadoresISEL.Controlador
{
    class WorkerProgressBar
    {

        public delegate void DelegateCRU(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion, string RFCpublico, string rfcOL, string rfcAnji);
        public delegate void DelegateOL(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion, string RFCpublico, string rfccru, string rfcmanuel);
        public delegate void DelegateISEL(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion, string RFCdario);
        public delegate void DelegateMANEUL(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion);

        public event DelegateCRU get_data_cru;
        public event DelegateOL get_data_ol;
        public event DelegateISEL get_data_isel;
        public event DelegateMANEUL get_data_manuel;

        public string fechainicial = "";
        public string fechafinal = "";
        public Controlador_Impresion controlaimpresion = null;
        public string RFCpublico = "";
        public string rfcOL="";
        public string rfcAnji = "";

        public void CRU_mtehod()
        {
            get_data_cru(fechainicial, fechafinal, controlaimpresion, RFCpublico, rfcOL, rfcAnji);
        }

        public string rfccru = "";
        public string rfcmanuel = "";

        public void OL_mtehod()
        {
            get_data_ol(fechainicial, fechafinal, controlaimpresion, RFCpublico, rfccru, rfcmanuel);
        }

        public string RFCdario = "";
        public void ISEL_mtehod()
        {
            get_data_isel(fechainicial, fechafinal, controlaimpresion, RFCdario);
        }



        public void MANUEL_mtehod()
        {
            get_data_manuel(fechainicial, fechafinal, controlaimpresion);
        }

    }
}
