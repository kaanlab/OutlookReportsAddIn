using System.Collections.Generic;
using System.ComponentModel;


namespace OutlookReportsAddIn.ViewModels
{
    public class BaseViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string proprtyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(proprtyName));
        }
    }
}
