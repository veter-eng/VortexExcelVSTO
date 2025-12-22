using CommunityToolkit.Mvvm.ComponentModel;

namespace VortexExcelAddIn.ViewModels
{
    /// <summary>
    /// Classe base para todos os ViewModels
    /// Herda de ObservableObject para implementar INotifyPropertyChanged
    /// </summary>
    public abstract class ViewModelBase : ObservableObject
    {
        // O CommunityToolkit.Mvvm.ComponentModel.ObservableObject já implementa:
        // - INotifyPropertyChanged
        // - SetProperty helper method
        // - OnPropertyChanged method

        // Métodos e propriedades comuns podem ser adicionados aqui
    }
}
