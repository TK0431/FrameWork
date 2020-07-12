using System;
using System.Windows.Input;

namespace FrameWork.ViewModels.Base
{
    public class RelayTCommand<T>:ICommand
    {
        private Action<T> mAction;

        public RelayTCommand(Action<T> action)
        {
            mAction = action;
        }

        public event EventHandler CanExecuteChanged = (sender, e) => { };

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            mAction((T)parameter);
        }
    }
}
