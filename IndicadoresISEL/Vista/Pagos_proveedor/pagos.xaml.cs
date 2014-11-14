using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using IndicadoresISEL.Controlador;
using IndicadoresISEL.Modelo;
using System.Windows.Forms;
using IndicadoresISEL.Vista.Cargador;
using System.Threading;

namespace IndicadoresISEL.Vista.Pagos_proveedor
{
    /// <summary>
    /// Lógica de interacción para pagos.xaml
    /// </summary>
    public partial class pagos 
    {

        Controlador__SDKAdmipaq controladorSDK;//para lalamr al controlador del sdk admipaq
        List<Tipos_Datos_CRU.FacturasCRU> ListDocmuentos;//liusta de todo los documentos
        Controlador_Impresion controlaimpresion;//para poder mandar a imprimir en PDF
        public pagos()
        {
            InitializeComponent();
            controladorSDK = new Controlador__SDKAdmipaq();
            controlaimpresion = new Controlador_Impresion();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            if (controladorSDK.GetConexion())//antes de hacer algo verifico si existe alguna conexion con alguna empresa
            {
                OnWorkerMethodStart();
            }
            else System.Windows.MessageBox.Show("Necesita Seleccionar una Empresa");//mando mensaje cuando no existe una empresa seleccionada
        }

        private void Selecciona_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                controladorSDK.ConexionAdmipaq(folderBrowserDialog1.SelectedPath);//Verifica si la carpeta seleccionada es correcta
                if (controladorSDK.GetConexion())
                {//como es correcta imprime en el label la direccion de la empresa
                    RuteEmpresa.Text = folderBrowserDialog1.SelectedPath + "\\";
                    //ahora ya tienes la empresa correctamente seleccionada
                }
            }
        }

        private void textBoxanio_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
            {
                e.Handled = false;
            }
            else { e.Handled = true; }
        }




        CargadorBar cargador;
        MainWindow ventaprincipal;
        private void OnWorkerMethodStart()
        {
            //creamos el objeto de nuestra clase 
            WorkerProgressBar workerfactura = new WorkerProgressBar();
            //por medio de un delegado instanciamos el metodo que se debera ejecutar en segundo plano, aqui seleccionamos el metodo logear
            //este metodo se encuentra en esta clase
            int mes = Convert.ToInt32(comboBoxMes.SelectedIndex) + 1;//obtengo el mes del cual se realizara el filtro
            string fechainicial = mes + "/01" + "/" + textBoxanio.Text;//obtengo mi fechaincial para mi filtro
            string fechafinal = mes + "/31" + "/" + textBoxanio.Text;//obtengo mi fecha final para mi filtro
            workerfactura.datos_ += new WorkerProgressBar.LogerDelegate(datos_);
            workerfactura.fechafinal = fechafinal; // le asignamos el correo a la clase creada (ingresado por el usuairo)
            workerfactura.fechainicial = fechainicial; // le asignamos el password a la clase creada (ingresado por el usuario)
            workerfactura.mes = mes.ToString();
            workerfactura.controlaimpresion = this.controlaimpresion;
            workerfactura.textBoxanio = textBoxanio.Text;
            workerfactura.RuteEmpresa = RuteEmpresa.Text;
            workerfactura.RFCpublico = RFCPublico.Text.Trim();
            workerfactura.rfc = "";
            //creamos el hilo para ejecutar el proceso en segundo plano, en el pasamos como argumento el metodo que queremos ejecutar
            //el metodo que se ejecutara es el metodo que se encuentra en la clase creado
            ThreadStart tStart = new ThreadStart(workerfactura.WorkerMethod);
            Thread t = new Thread(tStart); //iniciamos el hilo

            t.Start(); // inicializa el hilo

            cargador = new CargadorBar(); //Creamos el objeto de la clase CargadorBar (este clase contiene el cargador)
            cargador.Owner = ventaprincipal; //asignamos que este objeto es modela relacionando  cual es su propietario
            cargador.ShowDialog(); //mostramo el cargador (este metodo se ejecutara )

            //finalmente obtenemos el resultado del metodo logear para seleccionar la respuesta que tendra 


        }

        private void datos_(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion, string textBoxanio, string mes, string RuteEmpresa, string RFCpublico, string rfc)
        {
            List<Tipos_Datos_CRU.FacturasCRU> ListDocmuentos = new List<Tipos_Datos_CRU.FacturasCRU>();//inicializo mi lista donde tendramis documentos
            ListDocmuentos = controladorSDK.get_PagosProveedorCRU(fechainicial, fechafinal);//obtengo todas las listas de mis documentos conforme el filtro que se dio

            List<Tipos_Datos_CRU.FacturasCRU> list_rfc_publico = controladorSDK.FiltroRFCCRU(ListDocmuentos, RFCpublico);
            List<Tipos_Datos_CRU.FacturasCRU> list_rfc_ol = new List<Tipos_Datos_CRU.FacturasCRU>();


            controlaimpresion.ImpresionCRUPagosProveedor(ListDocmuentos, "01/" + mes + "/" + textBoxanio + "--" + "31/" + mes + "/" + textBoxanio, RuteEmpresa, list_rfc_publico);

            controlaimpresion.excel_import(ListDocmuentos, list_rfc_publico, list_rfc_ol, "Acumulado de abonos", "Abonos a público", "Abonos a OL");

            cargador.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Normal,
            new Action(
            delegate()
            {
                cargador.Close();
            }
            ));

        }

    }
}
