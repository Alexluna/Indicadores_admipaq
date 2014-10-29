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
using System.Windows.Shapes;
using Microsoft.Windows.Controls.Ribbon;
using IndicadoresISEL.Vista.Facturas;

namespace IndicadoresISEL.Vista.Contenedor_principal
{
    /// <summary>
    /// Interaction logic for Indicadores_pincipal.xaml
    /// </summary>
    public partial class Indicadores_pincipal : RibbonWindow
    {
        MainWindow mainwindow;//window de la principal
        public Indicadores_pincipal()
        {
            InitializeComponent();

            // Insert code required on object creation below this point.
        }

        #region get mainwindows
        /// <summary>
        /// metodo para obtener el main window principal
        /// </summary>
        /// <param name="mainwindow">obejto principal window</param>
        public void put_mainwindow(MainWindow mainwindow)
        {
            this.mainwindow = mainwindow;//obetengo el contenedor principal
        }
        #endregion


        #region evento closing ribbon
        /// <summary>
        /// metodo para cuando se va a cerrar el ribbon
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.mainwindow.Close();//cierro el mainwindow para que de esta forma mate todo el proyecto
        }
        #endregion

        private void Btnfacturacion_Click(object sender, RoutedEventArgs e)
        {
            facturacion factura = new facturacion();//crea el objeto del user contorl el cual se quiere mostrar
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(factura);
        }


    }
}
