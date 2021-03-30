using System.Windows;

namespace Invent
{
    /// <summary>
    /// Логика взаимодействия для SetValue.xaml
    /// </summary>
    public partial class SetValue : Window
    {
        public SetValue()
        {
            InitializeComponent();
        }

        private void Btn_setValue_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Txtbox_setValue.Text))
            {
                MessageBox.Show("Введите значение", "Установить значение", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            else
            {
                DialogResult = true;
            }
        }
    }
}
