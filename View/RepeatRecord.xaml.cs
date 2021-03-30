using System.Windows;

namespace Invent
{
    /// <summary>
    /// Логика взаимодействия для RepeatRecord.xaml
    /// </summary>
    public partial class RepeatRecord : Window
    {
        public RepeatRecord()
        {
            InitializeComponent();
        }

        private void Btn_accept_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
