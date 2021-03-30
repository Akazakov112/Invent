using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace Invent
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static MainWindow SelfRef { get; set; }

        /// <summary>
        /// Конструктор формы.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            SelfRef = this;
            // Если в каталоге с программой не существует папки InventWorks.
            if (!Directory.Exists(Environment.CurrentDirectory + @"\InventWorks\"))
            {
                // То создать ее.
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\InventWorks\");
            }
            // Подписка на события изменения базы данных ObservableCollection.
            OS.ScanList.CollectionChanged += ChangeCollection;
            // Подписка на событие загрузки формы с присваиванием scrollView таблицы сканирования дочернего элемента.
            Loaded += (o, e) => ForWorks.scrollViewer_dgScan = ForWorks.GetVisualChild<ScrollViewer>(Datagrid_scan);
            // Поулчить объект CollectionViewSource из ресурсов xaml.
            ForWorks.viewSource = (CollectionViewSource)TryFindResource("ScanCollection");
        }

        #region Методы

        /// <summary>
        /// Метод вызова окна выбора файла.
        /// </summary>
        private void ChooseInvFile()
        {
            // Инициализация окна выбора файла.
            var choosefileform = new ChooseFileForm
            {
                // Установка владельца дочернего окна.
                Owner = this
            };
            // Добавление в листбокс файлов из папки Invents внутри каталога программы.
            choosefileform.ListBox_chooseInvFile.ItemsSource = Directory.GetFiles(Environment.CurrentDirectory + @"\InventWorks\");
            // Если файлы в папке присутствуют, включить кнопку Выбрать файл.
            if (choosefileform.ListBox_chooseInvFile.Items.Count > 0) choosefileform.Btn_selectFile.IsEnabled = true;
            // Показать окно.
            choosefileform.ShowDialog();
        }

        /// <summary>
        /// Метод открытия файла и создания DataSet с таблицами данных.
        /// </summary>
        private async void OpenFileASync()
        {
            // Инициализация диалога для выбора файла.
            var ofd = new OpenFileDialog
            {
                // Настройка файл диалога.
                DefaultExt = "*.xlsx",
                Filter = "Excel 2007 + (*.xlsx)|*.xlsx",
                Title = "Выберите документ для загрузки данных"
            };
            // Если файл выбран.
            if (ofd.ShowDialog() == true)
            {
                // Прописать в лейбл статус.
                TxtBlock_importFileDesc.Text = "Идет загрузка файла";
                // Запустить бесконечный прогрессбар.
                ProgressBar_importDataFile.IsIndeterminate = true;
                // Если запущенный асинхронно метод импорта файла вернул истину.
                if (await Task.Run(() => OS.CreateDataBase(ofd.FileName, OS.DateBase)))
                {
                    // Вызвать метод выбора файла работ.
                    ChooseInvFile();
                    // Отключаем кнопку импорта файла.
                    Btn_importDataFile.IsEnabled = false;
                    MenuItem_addValue.IsEnabled = true;
                    MenuItem_delValue.IsEnabled = true;
                    MenuItem_setValue.IsEnabled = true;
                    MenuItem_searchValue.IsEnabled = true;
                    // Прописать в лейбл статус.
                    TxtBlock_importFileDesc.Text = string.Format("Файл загружен. Записей: {0}", OS.DateBase.Count);
                    // Остановить прогрессбар.
                    ProgressBar_importDataFile.IsIndeterminate = false;
                    // Интервал таймера 1 минута.
                    ForWorks.tempSaveTimer.Interval = 60000;
                    // Вызов события.
                    ForWorks.tempSaveTimer.Elapsed += TimerPick;
                    // Включение таймера.
                    ForWorks.tempSaveTimer.Enabled = true;
                    // Установить фокус на текстбокс для введения инвентарных номеров.
                    Txtbox_search.Focus();
                }
                // Если запущенный асинхронно метод импорта файла вернул ложь.
                else
                {
                    // Остановить прогрессбар.
                    ProgressBar_importDataFile.IsIndeterminate = false;
                    // Вывести в лейб статус.
                    TxtBlock_importFileDesc.Text = "Ошибка загрузки файла";
                }
            }
        }

        /// <summary>
        /// Метод определения параметров поискового запроса.
        /// </summary>
        private void SearchOs()
        {
            // Очищаем строки вывода текущей позиции.
            Txtblock_outInvNum.Text = string.Empty;
            Txtblock_outName.Text = string.Empty;
            Txtblock_outTypeOs.Text = string.Empty;
            Txtblock_outMarka.Text = string.Empty;
            Txtblock_outSn.Text = string.Empty;
            Txtblock_outDate.Text = string.Empty;
            Txtblock_outStatus.Text = string.Empty;
            Txtblock_outOtv.Text = string.Empty;
            Txtblock_outPodrazdelenie.Text = string.Empty;
            Txtblock_outOtdel.Text = string.Empty;
            // Если введенная строка пустая или состоит из пробелов.
            if (string.IsNullOrWhiteSpace(Txtbox_search.Text))
            {
                // Вывести ошибку.
                MessageBox.Show("Введите значение", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            //Иначе.
            else
            {
                // Переменная строки поиска с удаленными пробелами.
                string searchString = Txtbox_search.Text.Replace(" ", "");
                // Поиск по инвентарному номеру.
                OS searchItem = OS.DateBase.Find(x => x.InvNum == searchString);
                // Если объект поиска равен null.
                if (searchItem == null)
                {
                    // Поиск по серийному номеру.
                    searchItem = OS.DateBase.Find(x => x.SerialNum == searchString);
                }
                // Если после двух попыток поиска объект поиска не равен null. 
                if (searchItem != null)
                {
                    // Вывод данных по основному средству в лейблы вывода.
                    Txtblock_outInvNum.Text = searchItem.InvNum;
                    Txtblock_outName.Text = searchItem.Name;
                    Txtblock_outTypeOs.Text = searchItem.TypeOs;
                    Txtblock_outMarka.Text = searchItem.Marka;
                    Txtblock_outSn.Text = searchItem.SerialNum;
                    Txtblock_outDate.Text = searchItem.DatePostanovki;
                    Txtblock_outStatus.Text = searchItem.Sostoyanie;
                    Txtblock_outOtv.Text = searchItem.Otvetstvenniy;
                    Txtblock_outPodrazdelenie.Text = searchItem.Podrazdelenie;
                    Txtblock_outOtdel.Text = searchItem.Otdel;
                    // Если текстбоксы данных для ввода по умолчанию не пустые и не состоят из пробелов, 
                    // записывать данные в соответствующие совйства при формировании строки.
                    if (!string.IsNullOrWhiteSpace(Txtbox_location.Text))
                    {
                        searchItem.Location = Txtbox_location.Text;
                    }
                    if (!string.IsNullOrWhiteSpace(Txtbox_newOtv.Text))
                    {
                        searchItem.NewOtv = Txtbox_newOtv.Text;
                    }
                    if (!string.IsNullOrWhiteSpace(Txtbox_user.Text))
                    {
                        searchItem.NewUser = Txtbox_user.Text;
                    }
                    // Добавление записи в коллекцию отсканированных ОС.
                    if (OS.ScanList.Count > 0)
                    {
                        // Переменная для статуса повтора.
                        bool repeatStatus = false;
                        // Переменная для статуса добавления повтора.
                        bool repeatAdd = false;
                        // Перебираем элементы коллекции отсканированных объектов.
                        foreach (var item in OS.ScanList)
                        {
                            // Если нашлось совпадение по инвентарному номеру.
                            if (item.InvNum == searchItem.InvNum)
                            {
                                // Статус повтора изменить на истину.
                                repeatStatus = true;
                                // Текст информации о повторной записи.
                                var attentionDouble = string.Format(
                                    "Совпадение с позицией: {0}" +
                                    "\nОС: {1}" +
                                    "\nИнв. номер: {2}" +
                                    "\nМестоположение: {3}" +
                                    "\nдобавленной {4} позиций назад",
                                    item.NumRow, item.Name, item.InvNum, item.Location, OS.RowNumCounter - item.NumRow
                                    );
                                // Вывести окно запроса.
                                var window = new RepeatRecord
                                {
                                    Owner = this
                                };
                                // Установить текст повтора в окно.
                                window.Txtblock_sovpadenieDesc.Text = attentionDouble;
                                //Если будет нажата кнопка Записать.
                                if (window.ShowDialog() == true)
                                {
                                    // Переменной статуса добавления повтора присвоить истину.
                                    repeatAdd = true;
                                    // Установить значение свойства повторной строки на истину в объекте, с которым нашлось совпадение.
                                    item.RepeatRec = "Повтор";
                                }
                                // Выйти зи цикла.
                                break;
                            }
                        }

                        // Проверяем результаты проверки на повтор.
                        // Если статус добавления истина.
                        if (repeatAdd == true)
                        {
                            // Создать новый экземпляр с копированием значений свойств, кроме порядкогово номера.
                            OS repeatItem = new OS(searchItem)
                            {
                                // Присвоить порядковый номер строки.
                                NumRow = OS.RowNumCounter,
                                // Установить значение свойства повторной строки на истину в повторной строке.
                                RepeatRec = "Повтор"
                            };
                            // Добавить запись.
                            OS.ScanList.Add(repeatItem);
                            // Увеличить счетчик номеров строк.
                            OS.RowNumCounter++;
                        }
                        // Если повтора не обнаружено.
                        if (repeatStatus == false)
                        {
                            // Присвоить порядковый номер строки.
                            searchItem.NumRow = OS.RowNumCounter;
                            // Добавить запись.
                            OS.ScanList.Add(searchItem);
                            // Увеличить счетчик номеров строк.
                            OS.RowNumCounter++;
                        }
                    }
                    // Добавление первой записи в коллекцию, т.к. количество элементов коллекции равно 0.
                    else
                    {
                        // Присвоить порядковый номер строки.
                        searchItem.NumRow = OS.RowNumCounter;
                        // Добавить запись.
                        OS.ScanList.Add(searchItem);
                        // Увеличить счетчик номеров строк.
                        OS.RowNumCounter++;
                    }
                }
                // Если поиск не выполнен(состояние переменной ложно).
                else
                {
                    // Вывести ошибку.
                    MessageBox.Show("Не найдено!", "Поиск", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                // Очистить строку поиска.
                Txtbox_search.Clear();
                // Установить фокус на строку поиска.
                Txtbox_search.Focus();
            }
        }

        /// <summary>
        /// Обновляет и устаавливает фильтры в дефолт.
        /// </summary>
        public void SetFilters()
        {
            // Обновить фильтры.
            Filters.GetFilters();
            // Установить все фильтры на стандарт.
            ComboBox_selectOtv.SelectedItem = "Все";
            ComboBox_selectNewOtv.SelectedItem = "Все";
            ComboBox_selectDate.SelectedItem = "Все";
            ComboBox_selectLocation.SelectedItem = "Все";
            ComboBox_selectOtdel.SelectedItem = "Все";
            ComboBox_selectPodrazdelenie.SelectedItem = "Все";
            ComboBox_selectSostoyanie.SelectedItem = "Все";
            ComboBox_selectTypeOs.SelectedItem = "Все";
            ComboBox_selectUser.SelectedItem = "Все";
            ComboBox_selectTypeInvNum.SelectedItem = "Все";
            ComboBox_selectStatusRec.SelectedItem = "Все";
            ComboBox_selectColor.SelectedIndex = 0;
            CheckBox_repeatRec.IsChecked = false;
        }

        /// <summary>
        /// Изменяет свет ячейки в зависимости от нажатой кнопки.
        /// </summary>
        /// <param name="color"></param>
        private void SetColor(string color)
        {
            foreach (var t in Datagrid_scan.SelectedCells)
            {
                if (t.Item is OS item && item.RepeatRec != "Повтор")
                {
                    switch (t.Column.Header.ToString())
                    {
                        case "№":
                            item.NumRowColor = color;
                            break;
                        case "Инвентарный №":
                            item.InvNumColor = color;
                            break;
                        case "Наименование ОС":
                            item.NameColor = color;
                            break;
                        case "Тип ОС":
                            item.TypeOsColor = color;
                            break;
                        case "Марка":
                            item.MarkaColor = color;
                            break;
                        case "Серийный номер":
                            item.SNColor = color;
                            break;
                        case "Дата постановки на учет":
                            item.DateColor = color;
                            break;
                        case "Состояние":
                            item.SostoyanieColor = color;
                            break;
                        case "Ответственный":
                            item.OtvetstvenniyColor = color;
                            break;
                        case "Подразделение местоположения":
                            item.PodrazdelenieColor = color;
                            break;
                        case "Отдел местоположения":
                            item.OtdelColor = color;
                            break;
                        case "Комментарий":
                            item.CommentColor = color;
                            break;
                        case "Местоположение":
                            item.LocationColor = color;
                            break;
                        case "Новый пользователь":
                            item.NewUserColor = color;
                            break;
                        case "Пользователь":
                            item.UserColor = color;
                            break;
                        case "Новый ответственный":
                            item.NewOtvColor = color;
                            break;
                        case "Фактический SN":
                            item.FactSNColor = color;
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Выполняет поиск значения по свойствам объектов класса OS по поисковому запросу.
        /// </summary>
        /// <param name="props"></param>
        /// <param name="searchString"></param>
        private void SearchValue(List<PropertyInfo> props, string searchString, bool direction)
        {
            // Переменная ячейки датагрида.
            DataGridCellInfo dataGridCell;
            // Перебрать текущее представление датагрида.
            foreach (var temp in ForWorks.viewSource.View)
            {
                // Если элемент представления можно привести к типу OS.
                if (temp is OS item)
                {
                    // Перебрать коллекцию свойств типа OS.
                    foreach (var prop in props)
                    {
                        // Если значение свойства не null и оно совпадает с введенным значением в текстбокс окна поиска.
                        if (item.GetProperty(prop.Name).GetValue(item) != null && item.GetProperty(prop.Name).GetValue(item).ToString().ToLower().Contains(searchString.ToLower()))
                        {
                            // Перебрать все столбцы.
                            foreach (var col in Datagrid_scan.Columns)
                            {
                                // Найти совпадение привязанного к столбцу свойства и свойства объекта OS, в котором нашлось совпадение.
                                if (col.SortMemberPath == prop.Name)
                                {
                                    // Определить индекс столбца.
                                    int index = Datagrid_scan.Columns.IndexOf(col);
                                    // Присвоить переменной ячейки координаты найденной ячейки.
                                    dataGridCell = new DataGridCellInfo(item, Datagrid_scan.Columns[index]);
                                    // Если направление сверху вниз.
                                    if (direction)
                                    {
                                        // Перейти к проверке и выводу найденной ячейки.
                                        goto Found;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        // Метка выхода из циклов.
        Found:
            // Если переменная ячейки пустая.
            if (dataGridCell.Item == null)
            {
                MessageBox.Show("Не найдено", "Поиск значения", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            // Иначе/
            else
            {
                // Очистить коллекцию выделенных ячеек.
                Datagrid_scan.SelectedCells.Clear();
                // Установить ячейку текущей.
                Datagrid_scan.CurrentCell = dataGridCell;
            }
        }

        /// <summary>
        /// Вызывает окно поиска значения и передает выбранные параметры в метод поиска.
        /// </summary>
        private void OpenSearchForm()
        {
            // инициализация окна поиска.
            SearchValue window = new SearchValue();
            // Если нажата кнопка Найти.
            if (window.ShowDialog() == true)
            {
                if (window.Rbtn_searchInDoc.IsChecked == true)
                {
                    // Получить все свойства класса OS.
                    List<PropertyInfo> props = typeof(OS).GetProperties().ToList();
                    if (window.Rbtn_searchUp.IsChecked == true)
                    {
                        // вызвать метод поиска.
                        SearchValue(props, window.Txtbox_searchValue.Text, true);
                    }
                    else
                    {
                        // вызвать метод поиска.
                        SearchValue(props, window.Txtbox_searchValue.Text, false);
                    }

                }

                else if (window.Rbtn_searchInColumn.IsChecked == true)
                {
                    // Временная коллекция свойств из выбранных столбцов.
                    List<PropertyInfo> props = new List<PropertyInfo>();
                    // Перебираем выделенные ячейки.
                    foreach (var cell in Datagrid_scan.SelectedCells)
                    {
                        // Получаем привязанное к столбцу свойство и добавляем во временную коллекцию.
                        PropertyInfo test = typeof(OS).GetProperty(cell.Column.SortMemberPath);
                        props.Add(test);
                    }
                    if (window.Rbtn_searchUp.IsChecked == true)
                    {
                        // вызвать метод поиска.
                        SearchValue(props, window.Txtbox_searchValue.Text, true);
                    }
                    else
                    {
                        // вызвать метод поиска.
                        SearchValue(props, window.Txtbox_searchValue.Text, false);
                    }
                }
            }
        }

        /// <summary>
        /// Файл диалог для экспорта данных в Excel. Запускает метод сохранения асинхронно.
        /// </summary>
        private async void OpenFileDialogTOExcelASync()
        {
            // Диалог сохранения файла.
            SaveFileDialog sfd = new SaveFileDialog()
            {
                // Параметры диалога.
                DefaultExt = "*.xlsx",
                Filter = "Excel 2007 + |*.xlsx",
                Title = "Сохранение в Excel"
            };
            if (sfd.ShowDialog() == true)
            {
                Txtblock_ExpoerExcelDesc.Visibility = Visibility.Visible;
                ProgressBarExportToExcel.Visibility = Visibility.Visible;
                ProgressBarExportToExcel.IsIndeterminate = true;
                Btn_exportToExcel.IsEnabled = false;
                // Коллекция объектов текущего представления.
                ICollectionView view = ForWorks.viewSource.View;
                // Инициализация коллекции заголовков столбцов.
                List<string> colHeaders = new List<string>();
                // Записать все заголовки столбцов в коллекцию строк.
                foreach (var column in Datagrid_scan.Columns)
                {
                    colHeaders.Add(column.Header.ToString());
                }
                // При удачном сохранении вывести сообщение.
                if (await Task.Run(() => OS.SaveToExcel(sfd.FileName, colHeaders, view)))
                {
                    MessageBox.Show("Данные сохранены", "Сохранение в файл Excel", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                // При неудачном сохранении вывести ошибку.
                else
                {
                    MessageBox.Show("Ошибка сохранения!", "Сохранение в файл Excel", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                Txtblock_ExpoerExcelDesc.Visibility = Visibility.Hidden;
                ProgressBarExportToExcel.Visibility = Visibility.Hidden;
                ProgressBarExportToExcel.IsIndeterminate = false;
                Btn_exportToExcel.IsEnabled = true;
            }
        }

        #endregion

        #region Кнопки

        // Кнопка Загрузить файл данных.
        private void Btn_importDataFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileASync();
        }

        // Кнопка Поиск.
        private void Btn_startSearch_Click(object sender, RoutedEventArgs e)
        {
            SearchOs();
        }

        // Кнопка Очистить данные.
        private void Btn_clearScan_Click(object sender, RoutedEventArgs e)
        {
            // Запрос подтверждения удаления.
            MessageBoxResult checkBeforeDel = MessageBox.Show("Будут удалены все данные! Продолжить?", "Удаление данных", MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
            switch (checkBeforeDel)
            {
                // Если выбрано Да.
                case MessageBoxResult.Yes:
                    // Очищаем строки вывода текущей позиции.
                    Txtblock_outInvNum.Text = string.Empty;
                    Txtblock_outName.Text = string.Empty;
                    Txtblock_outTypeOs.Text = string.Empty;
                    Txtblock_outMarka.Text = string.Empty;
                    Txtblock_outSn.Text = string.Empty;
                    Txtblock_outDate.Text = string.Empty;
                    Txtblock_outStatus.Text = string.Empty;
                    Txtblock_outOtv.Text = string.Empty;
                    Txtblock_outPodrazdelenie.Text = string.Empty;
                    Txtblock_outOtdel.Text = string.Empty;
                    // Очистить переменную текущего файла работ.
                    ForWorks.currentWorkFile = string.Empty;
                    // Сбросить счетчик порядковых номеров строк.
                    OS.RowNumCounter = 1;
                    // Установка переменной состояния изменений документа в ложь.
                    ForWorks.checkEdit = false;
                    // Очистить список сканирования.
                    OS.ScanList.Clear();
                    // Сбросить пометки о состоянии записей в базе данных.
                    foreach (var obj in OS.DateBase)
                    {
                        obj.StatusRec = string.Empty;
                        obj.RepeatRec = string.Empty;
                    }
                    // Установить фокус на строку поиска.
                    Txtbox_search.Focus();
                    break;
                // Если выбрано Нет.
                case MessageBoxResult.No:
                    break;
            }
        }

        // Кнопка Сохранить.
        private void Btn_saveCurrentXmlFile_Click(object sender, RoutedEventArgs e)
        {
            // Если коллекция отсканированных элементов не пуская.
            if (OS.ScanList.Count > 0)
            {

                // Если сохранение успешно.
                if (OS.SaveXml(ForWorks.currentWorkFile))
                {
                    MessageBox.Show("Данные сохранены.", "Сохранить файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                // Если нет.
                else
                {
                    MessageBox.Show("Ошибка сохранения файла!", "Сохранить файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            // Если коллекция отсканированных элементов пуская.
            else
            {
                MessageBox.Show("Отсутствуют данные для сохранения.", "Сохранить файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // Кнопка Сохранить как.
        private void Btn_saveAsXmlFile_Click(object sender, RoutedEventArgs e)
        {
            // Если коллекция отсканированных элементов не пуская.
            if (OS.ScanList.Count > 0)
            {
                // Если сохранение успешно.
                if (OS.SaveXmlWithFileDialog())
                {
                    MessageBox.Show("Данные сохранены.", "Сохранить файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                // Если нет.
                else
                {
                    MessageBox.Show("Ошибка сохранения файла!", "Сохранить файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Отсутствуют данные для сохранения.", "Сохранить файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // Кнопка Открыть файл.
        private void Btn_openXmlFile_Click(object sender, RoutedEventArgs e)
        {
            // Если переменная статуса изменений в данных истинна(изменения были).
            if (ForWorks.checkEdit == true)
            {
                // Вызвать диалоговое окно с запросом сохранения данных.
                MessageBoxResult check = MessageBox.Show("Сохранить изменения в документе?", "Завершение работы", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
                // Если результат диалогового окна ДА.
                if (check == MessageBoxResult.Yes)
                {
                    // Если метод открытия файла вернет ложь.
                    if (OS.SaveXmlWithFileDialog())
                    {
                        // Если метод открытия файла вернет ложь.
                        if (!OS.OpenXmlWithFileDialogASync())
                        {
                            // Вывести ошибку.
                            MessageBox.Show("Ошибка открытия файла!", "Открыть файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
                // Если результат диалогового окна НЕТ.
                else if (check == MessageBoxResult.No)
                {
                    // Если метод открытия файла вернет ложь.
                    if (!OS.OpenXmlWithFileDialogASync())
                    {
                        // Вывести ошибку.
                        MessageBox.Show("Ошибка открытия файла!", "Открыть файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                // Если результат диалогового окна отмена.
                else
                {
                    return;
                }
            }
            // Если переменная статуса изменений в данных ложна(изменений не было).
            else
            {
                // Если метод открытия файла вернет ложь.
                if (!OS.OpenXmlWithFileDialogASync())
                {
                    // Вывести ошибку.
                    MessageBox.Show("Ошибка открытия файла!", "Открыть файл инвентаризации", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // Кнопка Запрет редактирования. 
        private void Btn_stopEditing_Click(object sender, RoutedEventArgs e)
        {
            // Изменить режим редактирования датагрида в зависимости от текущего значения.
            Datagrid_scan.IsReadOnly = Datagrid_scan.IsReadOnly != true;
        }

        // Кнопка О программе.
        private void Btn_aboutApp_Click(object sender, RoutedEventArgs e)
        {
            new AboutProg().ShowDialog();
        }

        // Кнопка Выбрать все(столбцы для отображения)
        private void Btn_checkAllColumns_Click(object sender, RoutedEventArgs e)
        {
            // Установить все чекбоксы отмеченными.
            InvNumColumnVisible.IsChecked = true;
            NameColumnVisible.IsChecked = true;
            TypeOsColumnVisible.IsChecked = true;
            MarkaColumnVisible.IsChecked = true;
            SNColumnVisible.IsChecked = true;
            DateColumnVisible.IsChecked = true;
            SostoyanieColumnVisible.IsChecked = true;
            OtvColumnVisible.IsChecked = true;
            PodrazdelenieColumnVisible.IsChecked = true;
            OtdelColumnVisible.IsChecked = true;
            CommentColumnVisible.IsChecked = true;
            LocationColumnVisible.IsChecked = true;
            UserColumnVisible.IsChecked = true;
            NewOtvColumnVisible.IsChecked = true;
            FactSnColumnVisible.IsChecked = true;
            StatusRecColumnVisible.IsChecked = true;
            // Установить все столбцы отображаемыми.
            InvNumCol.Visibility = Visibility.Visible;
            NameCol.Visibility = Visibility.Visible;
            TypeOsCol.Visibility = Visibility.Visible;
            MarkaCol.Visibility = Visibility.Visible;
            SNCol.Visibility = Visibility.Visible;
            DateCol.Visibility = Visibility.Visible;
            SostoyanieCol.Visibility = Visibility.Visible;
            OtvCol.Visibility = Visibility.Visible;
            PodrazdlenieCol.Visibility = Visibility.Visible;
            OtdelCol.Visibility = Visibility.Visible;
            CommentCol.Visibility = Visibility.Visible;
            LocationCol.Visibility = Visibility.Visible;
            UserCol.Visibility = Visibility.Visible;
            NewOtvCol.Visibility = Visibility.Visible;
            FactSnCol.Visibility = Visibility.Visible;
            StatusRecCol.Visibility = Visibility.Visible;
            // Обновить датагрид.
            Datagrid_scan.Items.Refresh();
        }

        // Кнопка Снять все(столбцы для отображения)
        private void Btn_uncheckAllColumns_Click(object sender, RoutedEventArgs e)
        {
            // Установить все чекбоксы неотмеченными.
            InvNumColumnVisible.IsChecked = false;
            NameColumnVisible.IsChecked = false;
            TypeOsColumnVisible.IsChecked = false;
            MarkaColumnVisible.IsChecked = false;
            SNColumnVisible.IsChecked = false;
            DateColumnVisible.IsChecked = false;
            SostoyanieColumnVisible.IsChecked = false;
            OtvColumnVisible.IsChecked = false;
            PodrazdelenieColumnVisible.IsChecked = false;
            OtdelColumnVisible.IsChecked = false;
            CommentColumnVisible.IsChecked = false;
            LocationColumnVisible.IsChecked = false;
            UserColumnVisible.IsChecked = false;
            NewOtvColumnVisible.IsChecked = false;
            FactSnColumnVisible.IsChecked = false;
            StatusRecColumnVisible.IsChecked = false;
            // Установить все столбцы скрытыми.
            InvNumCol.Visibility = Visibility.Collapsed;
            NameCol.Visibility = Visibility.Collapsed;
            TypeOsCol.Visibility = Visibility.Collapsed;
            MarkaCol.Visibility = Visibility.Collapsed;
            SNCol.Visibility = Visibility.Collapsed;
            DateCol.Visibility = Visibility.Collapsed;
            SostoyanieCol.Visibility = Visibility.Collapsed;
            OtvCol.Visibility = Visibility.Collapsed;
            PodrazdlenieCol.Visibility = Visibility.Collapsed;
            OtdelCol.Visibility = Visibility.Collapsed;
            CommentCol.Visibility = Visibility.Collapsed;
            LocationCol.Visibility = Visibility.Collapsed;
            UserCol.Visibility = Visibility.Collapsed;
            NewOtvCol.Visibility = Visibility.Collapsed;
            FactSnCol.Visibility = Visibility.Collapsed;
            StatusRecCol.Visibility = Visibility.Collapsed;
        }

        // Кнопка Применить фильтры.
        private void Btn_acceptFilters_Click(object sender, RoutedEventArgs e)
        {
            // Сохранить выбранные элементы комбобоксов фильтров.
            var tempotv = ComboBox_selectOtv.SelectedItem;
            var tempdate = ComboBox_selectDate.SelectedItem;
            var temploc = ComboBox_selectLocation.SelectedItem;
            var tempnewotv = ComboBox_selectNewOtv.SelectedItem;
            var tempotdel = ComboBox_selectOtdel.SelectedItem;
            var temppod = ComboBox_selectPodrazdelenie.SelectedItem;
            var tempsos = ComboBox_selectSostoyanie.SelectedItem;
            var tempstat = ComboBox_selectStatusRec.SelectedItem;
            var temptypeinv = ComboBox_selectTypeInvNum.SelectedItem;
            var temptypeos = ComboBox_selectTypeOs.SelectedItem;
            var tempuser = ComboBox_selectUser.SelectedItem;
            var tempcolor = ComboBox_selectColor.SelectedItem;
            // Подписаться на событие фильтрации элементов.
            ForWorks.viewSource.Filter += new FilterEventHandler(CollectionViewFilter);
            // Обновить фильтры.
            Filters.GetFilters();
            // Установить сохраненные выбранные элементы.
            ComboBox_selectOtv.SelectedItem = tempotv;
            ComboBox_selectDate.SelectedItem = tempdate;
            ComboBox_selectLocation.SelectedItem = temploc;
            ComboBox_selectNewOtv.SelectedItem = tempnewotv;
            ComboBox_selectOtdel.SelectedItem = tempotdel;
            ComboBox_selectPodrazdelenie.SelectedItem = temppod;
            ComboBox_selectSostoyanie.SelectedItem = tempsos;
            ComboBox_selectStatusRec.SelectedItem = tempstat;
            ComboBox_selectTypeInvNum.SelectedItem = temptypeinv;
            ComboBox_selectTypeOs.SelectedItem = temptypeos;
            ComboBox_selectUser.SelectedItem = tempuser;
            ComboBox_selectColor.SelectedItem = tempcolor;
        }

        // Кнопка Сбросить фильтры.
        private void Btn_clearFilters_Click(object sender, RoutedEventArgs e)
        {
            // Подписаться на событие фильтрации элементов.
            ForWorks.viewSource.Filter -= new FilterEventHandler(CollectionViewFilter);
            // Обновить фильтры.
            SetFilters();
        }

        // Кнопка красного цвета.
        private void Btn_red_Click(object sender, RoutedEventArgs e)
        {
            SetColor("Red");
        }

        // Кнопка оранжевого цвета.
        private void Btn_orange_Click(object sender, RoutedEventArgs e)
        {
            SetColor("Orange");
        }

        // Кнопка желтого цвета.
        private void Btn_yellow_Click(object sender, RoutedEventArgs e)
        {
            SetColor("Yellow");
        }

        // Кнопка зеленого цвета.
        private void Btn_green_Click(object sender, RoutedEventArgs e)
        {
            SetColor("Green");
        }

        // Кнопка синего цвета.
        private void Btn_blue_Click(object sender, RoutedEventArgs e)
        {
            SetColor("Blue");
        }

        // Кнопка фиолетового цвета.
        private void Btn_purple_Click(object sender, RoutedEventArgs e)
        {
            SetColor("Purple");
        }

        // Кнопка сброса цвета.
        private void Btn_resetColor_Click(object sender, RoutedEventArgs e)
        {
            SetColor("White");
        }

        // Кнопка Сохранить на второй вкладке.
        private void Btn_saveCurrentXmlDouble_Click(object sender, RoutedEventArgs e)
        {
            Btn_saveCurrentXmlFile_Click(sender, e);
        }

        // Кнопка Сохранить как на второй вкладке.
        private void Btn_saveAsXmlDouble_Click(object sender, RoutedEventArgs e)
        {
            Btn_saveAsXmlFile_Click(sender, e);
        }

        // Кнопка Выгрузить инвентарные номера на второй вкладке.
        private void Btn_exportInvNumToTxt_Click(object sender, RoutedEventArgs e)
        {
            OS.SaveInvNumToTxt();
        }

        // Кнопка Экспорт в файл Excel.
        private void Btn_exportToExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialogTOExcelASync();
        }

        #endregion

        #region События интерфейса

        // Событие для вызова таймером (сохранение промежуточных рещультатов).
        private void TimerPick(object timerTempSave, EventArgs Empty)
        {
            // Если в списке отсканированных ОС есть объекты.
            if (OS.ScanList.Count > 0 && ForWorks.checkEdit == true)
            {
                // Создание и запуск нового потока для выполнения фонового сохранения данных.
                var saveTemp = new Thread(new ParameterizedThreadStart(OS.SaveTempXml));
                saveTemp.Start(ForWorks.currentWorkFile);
            }
        }

        // Нажатие клавиши Enter в текстбоксах поиска.
        private void Txtbox_KeyDown(object sender, KeyEventArgs e)
        {
            // Если нажата клавиша Enter.
            if (e.Key == Key.Enter)
            {
                // Вызвать событие нажатия кнопки поиска.
                Btn_startSearch_Click(sender, e);
            }
        }

        // Событие изменения коллекции.
        public void ChangeCollection(object sender, NotifyCollectionChangedEventArgs e)
        {
            switch (e.Action)
            {
                // Если добавление.
                case NotifyCollectionChangedAction.Add:
                    // Обновить фильтры.
                    SetFilters();
                    // Установка переменной состояния изменений документа в истину.
                    ForWorks.checkEdit = true;
                    // Прокрутить датагрид вниз до конца.
                    ForWorks.scrollViewer_dgScan.ScrollToEnd();
                    break;
                // Если удаление.
                case NotifyCollectionChangedAction.Remove:
                    // Обновить фильтры.
                    SetFilters();
                    // Установка переменной состояния изменений документа в истину.
                    ForWorks.checkEdit = true;
                    break;
                // Если замена.
                case NotifyCollectionChangedAction.Replace:
                    // Обновить фильтры.
                    SetFilters();
                    // Установка переменной состояния изменений документа в истину.
                    ForWorks.checkEdit = true;
                    break;
            }
        }

        // Событие выбора элемента в комбобоксе строки отсканированного ОС.
        private void ComboboxRowSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Создаем переменную типа комбобокс путем преобразования объекта sender в комбобокс.
            var comboBox = sender as ComboBox;
            // Создаем переменную типа OS путем приведения типа выбранного итема датагрид.
            var osObj = (OS)Datagrid_scan.CurrentItem;
            // Если полученные объекты класса OS и комбобокса не являються объектами null.
            if (osObj != (object)null && comboBox != (object)null && comboBox.SelectedItem != null)
            {
                // Присвоить значение свойства цвета строки равному выбранному элементу в комбобокс.
                osObj.StatusRec = comboBox.SelectedItem.ToString();
                // Установка переменной состояния изменений документа в истину.
                ForWorks.checkEdit = true;
            }
        }

        // Событие завершения редактирования ячейки датагрида.
        private void Datagrid_scan_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            // Установка переменной состояния изменений документа в истину.
            ForWorks.checkEdit = true;
            // Обновить фильтры.
            SetFilters();
        }

        // Событие перед закрытием главного окна.
        private void MainForm_Closing(object sender, CancelEventArgs e)
        {
            e.Cancel = ForWorks.CheckSaveBeforeClosing();
        }

        // Обработчик события нажатия на чекбоксы отображения столбцов.
        private void ColumnVisible_Click(object sender, RoutedEventArgs e)
        {
            // Привести sender к типу чекбокс
            var Checkbox = (CheckBox)sender;
            // В зависимости от имени чекбокса изменить свойство отображения соответствующего столбца по состоянию чекбокса.
            switch (Checkbox.Name)
            {
                case "InvNumColumnVisible":
                    InvNumCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "NameColumnVisible":
                    NameCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "TypeOsColumnVisible":
                    TypeOsCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "MarkaColumnVisible":
                    MarkaCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "SNColumnVisible":
                    SNCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "DateColumnVisible":
                    DateCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "SostoyanieColumnVisible":
                    SostoyanieCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "OtvColumnVisible":
                    OtvCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "PodrazdelenieColumnVisible":
                    PodrazdlenieCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "OtdelColumnVisible":
                    OtdelCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "CommentColumnVisible":
                    CommentCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "LocationColumnVisible":
                    LocationCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "UserColumnVisible":
                    UserCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "NewOtvColumnVisible":
                    NewOtvCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "FactSnColumnVisible":
                    FactSnCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
                case "StatusRecColumnVisible":
                    StatusRecCol.Visibility = Checkbox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
                    break;
            }
        }

        // СОбытие фильтрации элементов коллекции отсканированных объектов.
        private void CollectionViewFilter(object sender, FilterEventArgs e)
        {
            if (e.Item is OS obj
                && ComboBox_selectDate.SelectedItem != null
                && ComboBox_selectLocation.SelectedItem != null
                && ComboBox_selectNewOtv.SelectedItem != null
                && ComboBox_selectOtdel.SelectedItem != null
                && ComboBox_selectOtv.SelectedItem != null
                && ComboBox_selectPodrazdelenie.SelectedItem != null
                && ComboBox_selectSostoyanie.SelectedItem != null
                && ComboBox_selectTypeInvNum.SelectedItem != null
                && ComboBox_selectTypeOs.SelectedItem != null
                && ComboBox_selectUser.SelectedItem != null)
            {
                bool IsAccept = true;

                IsAccept = IsAccept && (
                    ComboBox_selectOtv.SelectedItem.ToString() == "Все"
                    || ComboBox_selectOtv.SelectedItem.ToString() == "Пустые значения" && obj.Otvetstvenniy == ""
                    || obj.Otvetstvenniy == ComboBox_selectOtv.SelectedItem.ToString()
                    );

                IsAccept = IsAccept &&
                    (ComboBox_selectNewOtv.SelectedItem.ToString() == "Все"
                    || ComboBox_selectNewOtv.SelectedItem.ToString() == "Пустые значения" && obj.NewOtv == ""
                    || obj.NewOtv == ComboBox_selectNewOtv.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectUser.SelectedItem.ToString() == "Все"
                    || ComboBox_selectUser.SelectedItem.ToString() == "Пустые значения" && obj.User == ""
                    || obj.User == ComboBox_selectUser.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectLocation.SelectedItem.ToString() == "Все"
                    || ComboBox_selectLocation.SelectedItem.ToString() == "Пустые значения" && obj.Location == ""
                    || obj.Location == ComboBox_selectLocation.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectPodrazdelenie.SelectedItem.ToString() == "Все"
                    || ComboBox_selectPodrazdelenie.SelectedItem.ToString() == "Пустые значения" && obj.Podrazdelenie == ""
                    || obj.Podrazdelenie == ComboBox_selectPodrazdelenie.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectTypeOs.SelectedItem.ToString() == "Все"
                    || ComboBox_selectTypeOs.SelectedItem.ToString() == "Пустые значения" && obj.TypeOs == ""
                    || obj.TypeOs == ComboBox_selectTypeOs.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectSostoyanie.SelectedItem.ToString() == "Все"
                    || ComboBox_selectSostoyanie.SelectedItem.ToString() == "Пустые значения" && obj.Sostoyanie == ""
                    || obj.Sostoyanie == ComboBox_selectSostoyanie.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectTypeInvNum.SelectedItem.ToString() == "Все"
                    || ComboBox_selectTypeInvNum.SelectedItem.ToString() == "Деловые линии" && obj.InvNum.StartsWith("29600")
                    || ComboBox_selectTypeInvNum.SelectedItem.ToString() == "КЦ" && obj.InvNum.StartsWith("296CC")
                    || ComboBox_selectTypeInvNum.SelectedItem.ToString() == "ДЛ Транс" && obj.InvNum.StartsWith("6000")
                    || ComboBox_selectTypeInvNum.SelectedItem.ToString() == "гет Карго" && obj.InvNum.StartsWith("11100")
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectDate.SelectedItem.ToString() == "Все"
                    || ComboBox_selectDate.SelectedItem.ToString() == "Пустые значения" && obj.DatePostanovki == ""
                    || obj.DatePostanovki == ComboBox_selectDate.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectOtdel.SelectedItem.ToString() == "Все"
                    || ComboBox_selectOtdel.SelectedItem.ToString() == "Пустые значения" && obj.Otdel == ""
                    || obj.Otdel == ComboBox_selectOtdel.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    ComboBox_selectStatusRec.SelectedItem.ToString() == "Все"
                    || ComboBox_selectStatusRec.SelectedItem.ToString() == "Пустые значения" && obj.StatusRec == ""
                    || obj.StatusRec == ComboBox_selectStatusRec.SelectedItem.ToString()
                    );

                IsAccept = IsAccept && (
                    CheckBox_repeatRec.IsChecked == false
                    || CheckBox_repeatRec.IsChecked == true && obj.RepeatRec == "Повтор"
                    );

                if (ComboBox_selectColor.SelectedItem is TextBlock color)
                {
                    IsAccept = IsAccept && (
                    obj.NumRowColor == color.Text
                    || obj.InvNumColor == color.Text
                    || obj.NameColor == color.Text
                    || obj.TypeOsColor == color.Text
                    || obj.MarkaColor == color.Text
                    || obj.SNColor == color.Text
                    || obj.DateColor == color.Text
                    || obj.SostoyanieColor == color.Text
                    || obj.OtvetstvenniyColor == color.Text
                    || obj.PodrazdelenieColor == color.Text
                    || obj.OtdelColor == color.Text
                    || obj.CommentColor == color.Text
                    || obj.LocationColor == color.Text
                    || obj.UserColor == color.Text
                    || obj.NewOtvColor == color.Text
                    || obj.FactSNColor == color.Text
                    ); ;
                }

                e.Accepted = IsAccept;
            }
        }

        // Пункт контекстного меню - Добавить строку.
        private void MenuItem_addValue_Click(object sender, RoutedEventArgs e)
        {
            // Создать новый объекта класса основных средств с заполнением порядкового номера и коллекции статусов, остальные значения пустые. 
            var newitem = new OS(OS.RowNumCounter, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                OS.statusOs, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
            // Добавление объекта к коллекцию отсканироанных ОС.
            OS.ScanList.Add(newitem);
            // Увелические счетчика порядкового номера строк.
            OS.RowNumCounter++;
        }

        // Пункт контекстного меню - Удалить строку.
        private void MenuItem_delValue_Click(object sender, RoutedEventArgs e)
        {
            if (Datagrid_scan.SelectedCells.Count > 0)
            {
                // Временная коллекция выбранных объектов.
                List<OS> tempItems = new List<OS>();
                // Перебираем выделенные ячейки, приводим их в типу OS и добавляем во временную коллекцию.
                foreach (var cell in Datagrid_scan.SelectedCells)
                {
                    if (cell.Item is OS item) tempItems.Add(item);
                }
                // Перебираем временную коллекцию.
                foreach (var item in tempItems)
                {
                    // Если объект содержиться в коллекции, то удалить его по индексу.
                    if (OS.ScanList.Contains(item))
                    {
                        int indexItem = OS.ScanList.IndexOf(item);
                        OS.ScanList.RemoveAt(indexItem);
                    }
                }
            }
        }

        // Пункт контекстного меню - Установить значение.
        private void MenuItem_setValue_Click(object sender, RoutedEventArgs e)
        {
            // Инициализация окна установки значения.
            SetValue window = new SetValue();
            // Если нажата клавиза Записать.
            if (window.ShowDialog() == true && Datagrid_scan.SelectedCells.Count > 0)
            {
                // Перебираем выделенные ячейки, приводим их в типу OS и добавляем во временную коллекцию.
                foreach (var cell in Datagrid_scan.SelectedCells)
                {
                    // Если объект выбранной ячейки можно привести к типу OS, ячейки не принажделаж столбцу, 
                    //привязанному к свойству порядкового номера строки и не находятся в столбце с именем Статус ОС.
                    if (cell.Item is OS item && cell.Column.SortMemberPath != "NumRow" && cell.Column.Header.ToString() != "Статус ОС")
                    {
                        // Установить значение в свойство по привязке к столбцу.
                        item.GetProperty(cell.Column.SortMemberPath).SetValue(item, window.Txtbox_setValue.Text);
                        // Присвоить переменной состояния изменений документа истину.
                        ForWorks.checkEdit = true;
                    }
                }
            }
        }

        // Пункт контекстного меню - Найти.
        private void MenuItem_searchValue_Click(object sender, RoutedEventArgs e)
        {
            OpenSearchForm();
        }

        #endregion
    }
}