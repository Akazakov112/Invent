using System;
using System.IO;
using System.Windows;

namespace Invent
{
    /// <summary>
    /// Логика взаимодействия для ChooseFileForm.xaml
    /// </summary>
    public partial class ChooseFileForm : Window
    {
        public ChooseFileForm()
        {
            InitializeComponent();
        }
        
        // Кнопка Выбрать файл.
        private void Btn_selectFile_Click(object sender, RoutedEventArgs e)
        {
            // Если выбранный объект не является null.
            if (ListBox_chooseInvFile.SelectedItem != null)
            {
                if (OS.OpenXml(ListBox_chooseInvFile.SelectedItem.ToString()))
                {
                    // Закрыть окно.
                    Close();
                }
                else
                {
                    // Вывести ошибку.
                    MessageBox.Show("Ошибка открытия файла!", "Открыть файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                // Вывести ошибку.
                MessageBox.Show("Выберите файл!", "Открыть файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // Кнопка Создать новый файл.
        private void Btn_createNewFile_Click(object sender, RoutedEventArgs e)
        {
            // Путь нового файла.
            string filePath = Environment.CurrentDirectory + ForWorks.fileName;
            // Если сохранение новго файла успешно.
            if (OS.SaveXml(filePath))
            {
                // Установить счетчик строк на 1.
                OS.RowNumCounter = 1;
                // Закрыть окно.
                Close();
            }
            else
            {
                // Иначе вывести ошибку.
                MessageBox.Show("Ошибка создания файла!", "Создание файла инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Кнопка Обзор.
        private void Btn_findFile_Click(object sender, RoutedEventArgs e)
        {
            // Вызвать метод открытия файла.
            if (OS.OpenXmlWithFileDialogASync())
            {
                Close();
            }
            else
            {
                // Вывести ошибку.
                MessageBox.Show("Ошибка открытия файла!", "Открыть файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Кнопка Удалить файл.
        private void Btn_deleteFile_Click(object sender, RoutedEventArgs e)
        {
            // Если есть выбранный файл.
            if (ListBox_chooseInvFile.SelectedItems.Count > 0)
            {
                // Вызвать окно подтверждения.
                switch (MessageBox.Show("Вы уверены?", "Удаление файла", MessageBoxButton.YesNo, MessageBoxImage.Warning))
                {
                    // При выборе пункта Да.
                    case MessageBoxResult.Yes:
                        try
                        {
                            // Удалить файл.
                            File.Delete(ListBox_chooseInvFile.SelectedItem.ToString());
                            // Обновить список файлов в папке.
                            ListBox_chooseInvFile.ItemsSource = Directory.GetFiles(Environment.CurrentDirectory + @"\InventWorks\");
                        }
                        catch
                        {
                            // Иначе вывести ошибку.
                            MessageBox.Show("Ошибка удаления файла!", "Удаление файла инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        break;
                    // При выборе пункта Нет.
                    case MessageBoxResult.No:
                        break;
                }
            }
            // Иначе вывести сообщение.
            else
            {
                MessageBox.Show("Выберите файл", "Удаление файла", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
        }

        // Двойной клик по элементам листбокса
        private void ListBox_chooseInvFile_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Btn_selectFile_Click(sender, e);
        }
    }
}
