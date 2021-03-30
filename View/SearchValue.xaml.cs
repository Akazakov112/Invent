using System.Windows;

namespace Invent
{
    /// <summary>
    /// Логика взаимодействия для SearchValue.xaml
    /// </summary>
    public partial class SearchValue : Window
    {
        public SearchValue()
        {
            InitializeComponent();
        }

        private void Btn_searchValue_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
