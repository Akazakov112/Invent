using System;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace Invent
{
    /// <summary>
    /// Класс вспомогательных элементов.
    /// </summary>
    public class ForWorks
    {
        /// <summary>
        /// Переменная представления датагрида.
        /// </summary>
        public static CollectionViewSource viewSource;

        /// <summary>
        /// Переменная текущего файла.
        /// </summary>
        public static string currentWorkFile;

        /// <summary>
        /// Переменная состояния изменений документа.
        /// </summary>
        public static bool checkEdit = false;

        /// <summary>
        /// Имя для новых файлов инвертаризации.
        /// </summary>
        public static readonly string fileName = string.Format(@"\InventWorks\{0}.xml", DateTime.Now).Replace(":", "-");

        /// <summary>
        /// Поле объекта SCrollViewer для основного датагрида.
        /// </summary>
        public static ScrollViewer scrollViewer_dgScan;

        /// <summary>
        /// Таймер в 1 минуту для сохранения файла инвентаризации.
        /// </summary>
        public static readonly Timer tempSaveTimer = new Timer();


        /// <summary>
        /// Возвращает дочерний визуальный элемент от переданного родителя.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <returns></returns>
        public static T GetVisualChild<T>(DependencyObject parent) where T : Visual
        {
            T child = null;
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null) child = GetVisualChild<T>(v);
                if (child != null) break;
            }
            return child;
        }

        /// <summary>
        /// Выполняет проверку статуса изменений документа перед закрытием программы.
        /// </summary>
        /// <returns></returns>
        public static bool CheckSaveBeforeClosing()
        {
            // Если переменная статуса изменений в данных истинна(изменения были).
            if (checkEdit == true)
            {
                // Вызвать диалоговое окно с запросом сохранения данных.
                MessageBoxResult check = MessageBox.Show("Сохранить изменения в документе?", "Завершение работы", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
                // Если результат диалогового окна ДА.
                if (check == MessageBoxResult.Yes)
                {
                    // Если метод сохранения вернет истину.
                    if (OS.SaveXmlWithFileDialog())
                    {
                        // Вернуть ложь.
                        return false;
                    }
                    else
                    {
                        // Вернуть истину.
                        return true;
                    }
                }
                // Если результат диалогового окна НЕТ.
                else if (check == MessageBoxResult.No)
                {
                    // Вернуть ложь.
                    return false;
                }
                // Если результат диалогового окна отмена.
                else
                {
                    // Вернуть истину.
                    return true;
                }
            }
            // Если переменная статуса изменений в данных ложна(изменений не было).
            else
            {
                // Вернуть ложь.
                return false;
            }
        }
    }
}