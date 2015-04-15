using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace ImageFinder
{
    /// <summary>
    ///     Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private readonly StreamWriter _logFile = new StreamWriter("ImageFinder.log");

        private App()
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomainOnUnhandledException;
        }

        private void CurrentDomainOnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            _logFile.WriteLine(e.ExceptionObject);
            _logFile.WriteLine();
        }

        private void App_OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            _logFile.WriteLine(e.Exception);
            _logFile.WriteLine();
        }
    }
}