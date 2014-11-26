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
using IndicadoresISEL.Vista.Contenedor_principal;

namespace IndicadoresISEL
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            Indicadores_pincipal indicadores = new Indicadores_pincipal();/*creo mi obejto de mi ventana principal*/
            this.Visibility = Visibility.Collapsed;//collapso mi ventana de login (ventana de entrada)
            indicadores.put_mainwindow(this);//le mando el main window para poder cerrar la aplicación por competo
            indicadores.ShowDialog();//muestro mi ribbon 

        }


        /// <summary>
        /// boton el cual tiene la imagen de una X para poder cerrar el programa
        /// (para el programa)
        /// </summary>
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Boton para iniciar con el programa de indicadores
        /// </summary>
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Indicadores_pincipal indicadores = new Indicadores_pincipal();/*creo mi obejto de mi ventana principal*/
            this.Visibility = Visibility.Collapsed;//collapso mi ventana de login (ventana de entrada)
            indicadores.put_mainwindow(this);//le mando el main window para poder cerrar la aplicación por competo
            indicadores.ShowDialog();//muestro mi ribbon 
        }

       
    }
}
