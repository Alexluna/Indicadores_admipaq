using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IndicadoresISEL.Controlador
{
    class WorkerProgressBar
    {
        public delegate void LogerDelegate(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion, string textBoxanio, string mes, string RuteEmpresa, string RFCpublico, string rfc);
        public event LogerDelegate factura;

        public string fechainicial = "";
        public string fechafinal = "";
        public Controlador_Impresion controlaimpresion = null;
        public string textBoxanio="";
        public string mes = "";
        public string RuteEmpresa = "";
        public string RFCpublico = "";
        public string rfc = "";
        public void WorkerMethod()
        {
            factura(fechainicial, fechafinal, controlaimpresion, textBoxanio, mes, RuteEmpresa, RFCpublico, rfc);
        }
    }
}
