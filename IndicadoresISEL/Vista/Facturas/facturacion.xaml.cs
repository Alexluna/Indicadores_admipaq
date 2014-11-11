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
using IndicadoresISEL.Vista.Cargador;

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
            WorkerProgressBar myC = new WorkerProgressBar();
            //por medio de un delegado instanciamos el metodo que se debera ejecutar en segundo plano, aqui seleccionamos el metodo logear
            //este metodo se encuentra en esta clase
            int mes = Convert.ToInt32(comboBoxMes.SelectedIndex) + 1;//obtengo el mes del cual se realizara el filtro
            string fechainicial = mes + "/01" + "/" + textBoxanio.Text;//obtengo mi fechaincial para mi filtro
            string fechafinal = mes + "/31" + "/" + textBoxanio.Text;//obtengo mi fecha final para mi filtro
            myC.factura += new WorkerProgressBar.LogerDelegate(factura);
            myC.fechafinal = fechafinal; // le asignamos el correo a la clase creada (ingresado por el usuairo)
            myC.fechainicial = fechainicial; // le asignamos el password a la clase creada (ingresado por el usuario)
            myC.mes = mes.ToString();
            myC.controlaimpresion = this.controlaimpresion;
            myC.textBoxanio=textBoxanio.Text;
            myC.RuteEmpresa = RuteEmpresa.Text;
            myC.RFCpublico=RFCPublico.Text.Trim();
            myC.rfc=RFC.Text.Trim();
            //creamos el hilo para ejecutar el proceso en segundo plano, en el pasamos como argumento el metodo que queremos ejecutar
            //el metodo que se ejecutara es el metodo que se encuentra en la clase creado
            ThreadStart tStart = new ThreadStart(myC.WorkerMethod);
            Thread t = new Thread(tStart); //iniciamos el hilo

            t.Start(); // inicializa el hilo

            cargador = new CargadorBar(); //Creamos el objeto de la clase CargadorBar (este clase contiene el cargador)
            cargador.Owner = ventaprincipal; //asignamos que este objeto es modela relacionando  cual es su propietario
            cargador.ShowDialog(); //mostramo el cargador (este metodo se ejecutara )

            //finalmente obtenemos el resultado del metodo logear para seleccionar la respuesta que tendra 
            

        }

        private void factura(string fechainicial, string fechafinal, Controlador_Impresion controlaimpresion, string textBoxanio, string mes, string RuteEmpresa,string RFCpublico,string rfc)
        {
            List<Tipos_Datos_CRU.FacturasCRU>  ListDocmuentos = new List<Tipos_Datos_CRU.FacturasCRU>();//inicializo mi lista donde tendramis documentos
            ListDocmuentos = controladorSDK.get_Documentos(fechainicial, fechafinal);//obtengo todas las listas de mis documentos conforme el filtro que se dio

            List<Tipos_Datos_CRU.FacturasCRU> list_rfc_publico = controladorSDK.FiltroRFCCRU(ListDocmuentos, RFCpublico);
            List<Tipos_Datos_CRU.FacturasCRU> list_rfc_ol = controladorSDK.FiltroRFCCRU(ListDocmuentos, rfc);


            controlaimpresion.ImpresionCRUFacturas(ListDocmuentos, "01/" + mes + "/" + textBoxanio + "--" + "31/" + mes + "/" + textBoxanio, RuteEmpresa, list_rfc_publico, list_rfc_ol);

            controlaimpresion.excel_import(ListDocmuentos,list_rfc_publico,list_rfc_ol);

            cargador.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Normal,
            new Action(
            delegate()
            {
                cargador.Close();
            }
            ));

        }

        
        public void exportaExcel()
        {
            //fin título
            //worksheet.Cells[1, 1] = titulo;
            // storing header part in Excel
            //for (int i = 1; i < dataGrid1.Columns.Count + 1; i++)
            //{
            //    worksheet.Cells[2, i] = dataGrid1.Columns[i - 1].Header;
            //}



            //// storing Each row and column value to excel sheet
            //for (int i = 0; i < dataGrid1.Items.Count; i++)
            //{
            //    for (int j = 0; j < dataGrid1.Columns.Count; j++)
            //    {
            //        worksheet.Cells[i + 3, j + 1] = (dataGrid1.Items[i] as System.Data.DataRowView).Row.ItemArray[j].ToString();// .Cells[j].Value.ToString();
            //    }
            //}


            // save the application
            //workbook.SaveAs("c:\\output.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Exit from the application
            //app.Quit();
        }
       



        

          


       
       
    }
}
