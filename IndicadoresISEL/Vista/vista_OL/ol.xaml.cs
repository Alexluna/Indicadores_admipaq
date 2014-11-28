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
using System.Windows.Forms;
using IndicadoresISEL.Controlador;
using IndicadoresISEL.Vista.Cargador;
using System.Threading;
using IndicadoresISEL.Modelo;

namespace IndicadoresISEL.Vista.vista_OL
{
    /// <summary>
    /// Lógica de interacción para ol.xaml
    /// </summary>
    public partial class ol
    {
        Controlador__SDKAdmipaq controladorSDK;//para lalamr al controlador del sdk admipaq
        Controlador_Impresion controlaimpresion;//para poder mandar a imprimir en PDF
        public ol()
        {
            InitializeComponent();
            controladorSDK = new Controlador__SDKAdmipaq();//para manear y llamar metodos del controlador

            controlaimpresion = new Controlador_Impresion();


            //inicializo los datetime con la fecha actual
            dateinicial.SelectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            datefinal.SelectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
        }


        /// <summary>
        /// obtengo l a rua de la empresa y valido es es correcta
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            if (controladorSDK.GetConexion())//antes de hacer algo verifico si existe alguna conexion con alguna empresa
            {
                OnWorkerMethodStart();
            }
            else System.Windows.MessageBox.Show("Necesita Seleccionar una Empresa");//mando mensaje cuando no existe una empresa seleccionada
        }


        CargadorBar cargador;
        MainWindow ventaprincipal;
        private void OnWorkerMethodStart()
        {
            //creamos el objeto de nuestra clase 
            WorkerProgressBar workerfactura = new WorkerProgressBar();
            //por medio de un delegado instanciamos el metodo que se debera ejecutar en segundo plano, aqui seleccionamos el metodo logear
            //este metodo se encuentra en esta clase
            //int mes = Convert.ToInt32(comboBoxMes.SelectedIndex) + 1;//obtengo el mes del cual se realizara el filtro
            string fechainicial = dateinicial.SelectedDate.Value.Date.ToString("MM/dd/yyyy");//obtengo mi fechaincial para mi filtro
            string fechafinal = datefinal.SelectedDate.Value.Date.ToString("MM/dd/yyyy");//obtengo mi fecha final para mi filtro


            workerfactura.get_data_ol += new WorkerProgressBar.DelegateOL(get_data_);
            workerfactura.fechafinal = fechafinal; // le asignamos el correo a la clase creada (ingresado por el usuairo)
            workerfactura.fechainicial = fechainicial; // le asignamos el password a la clase creada (ingresado por el usuario)
            workerfactura.controlaimpresion = this.controlaimpresion;
            workerfactura.RFCpublico = RFCPublico.Text.Trim();
            workerfactura.rfccru = RfCRU.Text.Trim();
            workerfactura.rfcmanuel = RFCmanuel.Text.Trim();
            workerfactura.RFCpublico = RFCPublico.Text.Trim();

            //creamos el hilo para ejecutar el proceso en segundo plano, en el pasamos como argumento el metodo que queremos ejecutar
            //el metodo que se ejecutara es el metodo que se encuentra en la clase creado
            ThreadStart tStart = new ThreadStart(workerfactura.OL_mtehod);
            Thread t = new Thread(tStart); //iniciamos el hilo

            t.Start(); // inicializa el hilo

            cargador = new CargadorBar(); //Creamos el objeto de la clase CargadorBar (este clase contiene el cargador)
            cargador.Owner = ventaprincipal; //asignamos que este objeto es modela relacionando  cual es su propietario
            cargador.ShowDialog(); //mostramo el cargador (este metodo se ejecutara )

            //finalmente obtenemos el resultado del metodo logear para seleccionar la respuesta que tendra 


        }

        private void get_data_(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion, string RFCpublico, string rfccru, string rfcmanuel)
        {
            List<Tipos_Datos_CRU.CRU> ListDocmuentos = new List<Tipos_Datos_CRU.CRU>();//inicializo mi lista donde tendramis documentos
            ListDocmuentos = controladorSDK.get_Documentos_ol(fechainicial, fechafinal);//obtengo todas las listas de mis documentos conforme el filtro que se dio
            //debo de guardar cada uno en su propio objeto
            Tipos_Datos_CRU.ListDatosOL ListIndicadorres = controladorSDK.filtro_indicadores_ol_tipo(ListDocmuentos, RFCpublico, rfccru, rfcmanuel);
            controlaimpresion.excel_importOL(ListIndicadorres);
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
