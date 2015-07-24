using System;
using System.Windows;
using System.Windows.Input;

namespace GpfMeter.Commands
{
    public class ShowWindowCommand : ICommand
    {
        public void Execute(object parameter)
        {
            var w = (GpfMeter.MainWindow)parameter;
            w.StartGpfMeter();
        }


        public bool CanExecute(object parameter)
        {
            return parameter is Window;
        }
        // Ignored
        public event EventHandler CanExecuteChanged
        {
            add { }
            remove { }
        }
    }
}
