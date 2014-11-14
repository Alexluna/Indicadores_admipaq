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
using IndicadoresISEL.Vista.Abonos;
using IndicadoresISEL.Vista.compras;
using IndicadoresISEL.Vista.Pagos_proveedor;

namespace IndicadoresISEL.Vista.Contenedor_principal
{
    /// <summary>
    /// Interaction logic for Indicadores_pincipal.xaml
    /// </summary>
    public partial class Indicadores_pincipal : RibbonWindow
    {
        MainWindow mainwindow;//window de la principal
        facturacion factura;
        abono abonos;
        compra compras;
        pagos pagos_proveedor;
        public Indicadores_pincipal()
        {
            InitializeComponent();
            factura = new facturacion();
            abonos = new abono();
            compras = new compra();
            pagos_proveedor = new pagos();
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


        #region CRU
        /// <summary>
        /// click para user control de facturacion
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnfacturacion_Click(object sender, RoutedEventArgs e)
        {
            MarcarCasilla(Btnfacturacion);
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(factura);
        }


        /// <summary>
        /// click de abonos
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnabonos_Click(object sender, RoutedEventArgs e)
        {
            MarcarCasilla(Btnabonos);
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(abonos);
        }

        /// <summary>
        /// click de compras
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btncompras_Click(object sender, RoutedEventArgs e)
        {
            MarcarCasilla(Btncompras);
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(compras);
        }

        /// <summary>
        /// click de pagos al proveedor
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnpagos_proveedor_Click(object sender, RoutedEventArgs e)
        {
            MarcarCasilla(Btnpagos_proveedor);
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(pagos_proveedor);
        }
        #endregion



        #region OL
        /// <summary>
        /// click para user control de facturacion
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnfacturacion_ol_Click(object sender, RoutedEventArgs e)
        {
            MarcarCasilla(Btnfacturacion_ol);
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(factura);
        }


        /// <summary>
        /// click de abonos
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnabonos_ol_Click(object sender, RoutedEventArgs e)
        {
            MarcarCasilla(Btnabonos_ol);
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(abonos);
        }

        /// <summary>
        /// click de compras
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btncompras_ol_Click(object sender, RoutedEventArgs e)
        {
            MarcarCasilla(Btncompras_ol);
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(compras);
        }

        /// <summary>
        /// click de pagos al proveedor
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnpagos_proveedor_ol_Click(object sender, RoutedEventArgs e)
        {
            MarcarCasilla(Btnpagos_proveedor_ol);
            gridvista.Children.Clear();//limpio el contenido del grid
            gridvista.Children.Add(pagos_proveedor);
        }
        #endregion




        #region GET SELECT
        RibbonButton buttom;
        private LinearGradientBrush ObtenerDegradado()
        {
            LinearGradientBrush myLinearGradientBrush = new LinearGradientBrush();
            myLinearGradientBrush.StartPoint = new System.Windows.Point(0.5, 0);
            myLinearGradientBrush.EndPoint = new System.Windows.Point(0.5, 1);
            myLinearGradientBrush.GradientStops.Add(
                new GradientStop(System.Windows.Media.Color.FromRgb(253, 229, 202), 0.0));

            myLinearGradientBrush.GradientStops.Add(
                new GradientStop(System.Windows.Media.Color.FromRgb(255, 210, 118), 0.75));

            return myLinearGradientBrush;
        }
        private void MarcarCasilla(RibbonButton btn)
        {
            if (buttom == null)
            {
                buttom = btn;
                buttom.Background = ObtenerDegradado();
            }
            else
            {
                buttom.Background = null;
                buttom = btn;
                buttom.Background = ObtenerDegradado();
            }

        }
        #endregion

    }
}
