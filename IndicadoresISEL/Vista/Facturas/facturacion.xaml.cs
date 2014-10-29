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
using IndicadoresISEL.Modelo;
using System.Threading;
using System.Windows.Threading;
using System.ComponentModel;

namespace IndicadoresISEL.Vista.Facturas
{
    /// <summary>
    /// Lógica de interacción para facturacion.xaml
    /// </summary>
    public partial class facturacion 
    {
        Controlador__SDKAdmipaq controladorSDK;//para lalamr al controlador del sdk admipaq
        List<Tipos_Datos_CRU.FacturasCRU> ListDocmuentos;//liusta de todo los documentos
        Controlador_Impresion controlaimpresion;//para poder mandar a imprimir en PDF
        public facturacion()
        {
            InitializeComponent();
            controladorSDK = new Controlador__SDKAdmipaq();//para manear y llamar metodos del controlador
            ListDocmuentos = new List<Tipos_Datos_CRU.FacturasCRU>();
            controlaimpresion = new Controlador_Impresion();
        }

        /// <summary>
        /// Método para seleccionar una ruta de alguna empresa
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

        
       
        BackgroundWorker bw;
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            //String pp = "25.55";
            //System.Windows.MessageBox.Show(pp);
            //System.Windows.MessageBox.Show(float.Parse(pp, CultureInfo.InvariantCulture.NumberFormat)+"");
            if (controladorSDK.GetConexion())//antes de hacer algo verifico si existe alguna conexion con alguna empresa
            {
               
                load.Visibility = Visibility.Visible;
                
              /*  Action  EmptyDelegate = delegate() { };
                load.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate);*/

               /* bw = new BackgroundWorker();

                // Queremos estar informados del estado de la operación.
                // Esto nos habilita para poder llamar a bw.ReportProgress().
                bw.WorkerReportsProgress = true;

                // Queremos tener la posibilidad de cancelar el proceso.
                // Esto nos habilita para poder llamar a bw.CancelAsync().
                bw.WorkerSupportsCancellation = true;

                // Suscripción a eventos.
                

                // DoWork es el hilo en donde se va a realizar la operación.
                bw.DoWork += bw_DoWork;

                // Este evento es el que tenemos que aprovechar para ser informados
                // del estado de la operación.
                //bw.ProgressChanged += bw_ProgressChanged;

                // Este evento es el que se va a disparar cuando la tarea haya finalizado.
                // Aquí es donde vamos a poder recoger los resultados de la operación.
                bw.RunWorkerCompleted += bw_RunWorkerCompleted;

                // Una vez que ya hemos configurado el backgroundWorker como hemos querido
                // este método sirve para poner en marcha la operación
                // que potencialmente puede llevar un tiempo más o menos grande en realizarse.
                bw.RunWorkerAsync(this);*/
                ListDocmuentos = new List<Tipos_Datos_CRU.FacturasCRU>();//inicializo mi lista donde tendramis documentos
                int mes = Convert.ToInt32(comboBoxMes.SelectedIndex) + 1;//obtengo el mes del cual se realizara el filtro
                string fechainicial = mes + "/01" + "/" + textBoxanio.Text;//obtengo mi fechaincial para mi filtro
                string fechafinal = mes + "/31" + "/" + textBoxanio.Text;//obtengo mi fecha final para mi filtro
                ListDocmuentos = controladorSDK.get_Documentos(fechainicial, fechafinal);//obtengo todas las listas de mis documentos conforme el filtro que se dio
                ListDocmuentos = controladorSDK.get_Documentos(fechainicial, fechafinal);//obtengo todas las listas de mis documentos conforme el filtro que se dio

                //manda a imprimir al pdf
                //System.Windows.MessageBox.Show("PDF");
                //manda a llmar 2 eventos para realizar filtros por RFC (publico y ol) y despues los manda para que se impriman en un PDF
                //se envia la ruta de la empresa ya que en la ruta de la empresa se guardara el pdf y se manda el rango de fechas que es por mes para imprimirlo en el pdf
                controlaimpresion.ImpresionCRUFacturas(ListDocmuentos, "01/" + mes + "/" + textBoxanio.Text + "--" + "31/" + mes + "/" + textBoxanio.Text, RuteEmpresa.Text, controladorSDK.FiltroRFCCRU(ListDocmuentos, "XAXX010101000       "), controladorSDK.FiltroRFCCRU(ListDocmuentos, "OLU120912UM0        "));

                load.Visibility = Visibility.Hidden;
                
            }
            else System.Windows.MessageBox.Show("Necesita Seleccionar una Empresa");//mando mensaje cuando no existe una empresa seleccionada
        }



        

          


       
       
    }
}
