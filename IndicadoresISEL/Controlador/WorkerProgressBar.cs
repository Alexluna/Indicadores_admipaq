using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IndicadoresISEL.Controlador
{
    class WorkerProgressBar
    {
        
        public delegate void DelegateCRU(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion, string RFCpublico, string rfcOL, string rfcAnji);
        
        public event DelegateCRU get_data;

        public string fechainicial = "";
        public string fechafinal = "";
        public Controlador_Impresion controlaimpresion = null;
        public string RFCpublico = "";
        public string rfcOL="";
        public string rfcAnji = "";
        
        public void CRU_mtehod()
        {
            get_data(fechainicial,fechafinal,controlaimpresion,RFCpublico,rfcOL,rfcAnji);
        }

    }
}
