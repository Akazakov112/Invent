using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Xml.Serialization;
using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.Win32;

namespace Invent
{
    [Serializable]
    public class OS : INotifyPropertyChanged
    {
        #region Статические поля

        /// <summary>
        /// Коллекция значений статусов записей.
        /// </summary>
        public static readonly List<string> statusOs = new List<string>() { "", "Замена этикетки", "В резерв", "Замена этикетки и в резерв" };

        #endregion

        #region Статические свойства

        /// <summary>
        /// Коллекция базы данных всех ОС, выгружаемая из файла Excel.
        /// </summary>
        public static List<OS> DateBase { get; set; } = new List<OS>();

        /// <summary>
        /// Коллекция объектов класса для сканируемых ОС.
        /// </summary>
        public static ObservableCollection<OS> ScanList { get; set; } = new ObservableCollection<OS>();

        #endregion

        #region Поля

        private int _numrow;
        private string _invnum;
        private string _name;
        private string _typeos;
        private string _marka;
        private string _serialnum;
        private string _datepostanovki;
        private string _sostoyanie;
        private string _otvetstvenniy;
        private string _podrazdelenie;
        private string _otdel;
        private string _repeat;
        private string _comment;
        private string _location;
        private string _user;
        private string _newuser;
        private string _newotv;
        private string _factserialnum;
        private string _statusrec;

        private string _numrowcolor = "White";
        private string _invnumcolor = "White";
        private string _namecolor = "White";
        private string _typeoscolor = "White";
        private string _markacolor = "White";
        private string _serialnumcolor = "White";
        private string _datepostavkicolor = "White";
        private string _sostoyaniecolor = "White";
        private string _otvetstvenniycolor = "White";
        private string _podrazdeleniecolor = "White";
        private string _otdelcolor = "White";
        private string _commentcolor = "White";
        private string _locationcolor = "White";
        private string _usercolor = "White";
        private string _newusercolor = "White";
        private string _newotvcolor = "White";
        private string _factserialnumcolor = "White";

        private List<string> _statusRecList;

        #endregion

        #region Cвойства

        /// <summary>
        /// Счетчик порядкового номера строк.
        /// </summary>
        public static int RowNumCounter { get; set; }

        /// <summary>
        /// Порядковый номер объекта.
        /// </summary>
        public int NumRow
        {
            get { return _numrow; }
            set { _numrow = value; OnPropertyChanged("NumRow"); }
        }

        /// <summary>
        /// Инвентарный номер.
        /// </summary>
        public string InvNum
        {
            get { return _invnum; }
            set { _invnum = value; OnPropertyChanged("InvNum"); }

        }

        /// <summary>
        /// Наименование ОС.
        /// </summary>
        public string Name
        {
            get { return _name; }
            set { _name = value; OnPropertyChanged("Name"); }
        }

        /// <summary>
        /// Тип ОС.
        /// </summary>
        public string TypeOs
        {
            get { return _typeos; }
            set { _typeos = value; OnPropertyChanged("TypeOs"); }
        }

        /// <summary>
        /// Марка ОС.
        /// </summary>
        public string Marka
        {
            get { return _marka; }
            set { _marka = value; OnPropertyChanged("Marka"); }
        }

        /// <summary>
        /// Серийный номер.
        /// </summary>
        public string SerialNum
        {
            get { return _serialnum; }
            set { _serialnum = value; OnPropertyChanged("SerialNum"); }
        }

        /// <summary>
        /// Дата постановки на учет.
        /// </summary>
        public string DatePostanovki
        {
            get { return _datepostanovki; }
            set { _datepostanovki = value; OnPropertyChanged("DatePostanovki"); }
        }

        /// <summary>
        /// Состояние ОС.
        /// </summary>
        public string Sostoyanie
        {
            get { return _sostoyanie; }
            set { _sostoyanie = value; OnPropertyChanged("Sostoyanie"); }
        }

        /// <summary>
        /// Ответственный за ОС.
        /// </summary>
        public string Otvetstvenniy
        {
            get { return _otvetstvenniy; }
            set { _otvetstvenniy = value; OnPropertyChanged("Otvetstvenniy"); }
        }

        /// <summary>
        /// Подразделение учета.
        /// </summary>
        public string Podrazdelenie
        {
            get { return _podrazdelenie; }
            set { _podrazdelenie = value; OnPropertyChanged("Podrazdelenie"); }
        }

        /// <summary>
        /// Отдел подразделения учета.
        /// </summary>
        public string Otdel
        {
            get { return _otdel; }
            set { _otdel = value; OnPropertyChanged("Otdel"); }
        }

        /// <summary>
        /// Комментарий.
        /// </summary>
        public string Comment
        {
            get { return _comment; }
            set { _comment = value; OnPropertyChanged("Comment"); }
        }

        /// <summary>
        /// Местоположение.
        /// </summary>
        public string Location
        {
            get { return _location; }
            set { _location = value; OnPropertyChanged("Location"); }
        }

        /// <summary>
        /// Пользователь ОС.
        /// </summary>
        public string User
        {
            get { return _user; }
            set { _user = value; OnPropertyChanged("User"); }
        }

        /// <summary>
        /// Новый пользователь ОС.
        /// </summary>
        public string NewUser
        {
            get { return _newuser; }
            set { _newuser = value; OnPropertyChanged("NewUser"); }
        }

        /// <summary>
        /// Новый ответственный.
        /// </summary>
        public string NewOtv
        {
            get { return _newotv; }
            set { _newotv = value; OnPropertyChanged("NewOtv"); }
        }

        /// <summary>
        /// Фактический серийный номер.
        /// </summary>
        public string FactSerialNum
        {
            get { return _factserialnum; }
            set { _factserialnum = value; OnPropertyChanged("FactSerialNum"); }
        }

        /// <summary>
        /// Статус записи для DataGrid.
        /// </summary>
        public string StatusRec
        {
            get { return _statusrec; }
            set { _statusrec = value; OnPropertyChanged("StatusRec"); }
        }

        /// <summary>
        /// Значение, указывающее, является ли запись повторной.
        /// </summary>
        public string RepeatRec
        {
            get { return _repeat; }
            set { _repeat = value; OnPropertyChanged("RepeatRec"); }
        }

        /// <summary>
        /// Коллекция статусов записи.
        /// </summary>
        public List<string> StatusRecList
        {
            get { return _statusRecList; }
            set { _statusRecList = value; OnPropertyChanged("StatusRecList"); }
        }

        /// <summary>
        /// Цвет ячеек номера строки.
        /// </summary>
        public string NumRowColor
        {
            get { return _numrowcolor; }
            set { _numrowcolor = value; OnPropertyChanged("NumRowColor"); }
        }

        /// <summary>
        /// Цвет ячеек инвентарного номера.
        /// </summary>
        public string InvNumColor
        {
            get { return _invnumcolor; }
            set { _invnumcolor = value; OnPropertyChanged("InvNumColor"); }
        }

        /// <summary>
        /// Цвет ячеек наименования.
        /// </summary>
        public string NameColor
        {
            get { return _namecolor; }
            set { _namecolor = value; OnPropertyChanged("NameColor"); }
        }

        /// <summary>
        /// Цвет ячеек типа ОС.
        /// </summary>
        public string TypeOsColor
        {
            get { return _typeoscolor; }
            set { _typeoscolor = value; OnPropertyChanged("TypeOsColor"); }
        }

        /// <summary>
        /// Цвет ячеек марки.
        /// </summary>
        public string MarkaColor
        {
            get { return _markacolor; }
            set { _markacolor = value; OnPropertyChanged("MarkaColor"); }
        }

        /// <summary>
        /// Цвет ячеек серийного номера.
        /// </summary>
        public string SNColor
        {
            get { return _serialnumcolor; }
            set { _serialnumcolor = value; OnPropertyChanged("SNColor"); }
        }

        /// <summary>
        /// Цвет ячеек даты постановки.
        /// </summary>
        public string DateColor
        {
            get { return _datepostavkicolor; }
            set { _datepostavkicolor = value; OnPropertyChanged("DateColor"); }
        }

        /// <summary>
        /// Цвет ячеек состояния.
        /// </summary>
        public string SostoyanieColor
        {
            get { return _sostoyaniecolor; }
            set { _sostoyaniecolor = value; OnPropertyChanged("SostoyanieColor"); }
        }

        /// <summary>
        /// Цвет ячеек ответственного.
        /// </summary>
        public string OtvetstvenniyColor
        {
            get { return _otvetstvenniycolor; }
            set { _otvetstvenniycolor = value; OnPropertyChanged("OtvetstvenniyColor"); }
        }

        /// <summary>
        /// Цвет ячеек подразделения.
        /// </summary>
        public string PodrazdelenieColor
        {
            get { return _podrazdeleniecolor; }
            set { _podrazdeleniecolor = value; OnPropertyChanged("PodrazdelenieColor"); }
        }

        /// <summary>
        /// Цвет ячеек отдела.
        /// </summary>
        public string OtdelColor
        {
            get { return _otdelcolor; }
            set { _otdelcolor = value; OnPropertyChanged("OtdelColor"); }
        }

        /// <summary>
        /// Цвет ячеек комментария.
        /// </summary>
        public string CommentColor
        {
            get { return _commentcolor; }
            set { _commentcolor = value; OnPropertyChanged("CommentColor"); }
        }

        /// <summary>
        /// Цвет ячеек местоположения.
        /// </summary>
        public string LocationColor
        {
            get { return _locationcolor; }
            set { _locationcolor = value; OnPropertyChanged("LocationColor"); }
        }

        /// <summary>
        /// Цвет ячеек пользователя.
        /// </summary>
        public string UserColor
        {
            get { return _usercolor; }
            set { _usercolor = value; OnPropertyChanged("UserColor"); }
        }

        /// <summary>
        /// Цвет ячеек пользователя.
        /// </summary>
        public string NewUserColor
        {
            get { return _newusercolor; }
            set { _newusercolor = value; OnPropertyChanged("NewUserColor"); }
        }

        /// <summary>
        /// Цвет ячеек нового ответственного.
        /// </summary>
        public string NewOtvColor
        {
            get { return _newotvcolor; }
            set { _newotvcolor = value; OnPropertyChanged("NewOtvColor"); }
        }

        /// <summary>
        /// Цвет ячеек фактического инвентарного номера.
        /// </summary>
        public string FactSNColor
        {
            get { return _factserialnumcolor; }
            set { _factserialnumcolor = value; OnPropertyChanged("FactSNColor"); }
        }

        #endregion

        #region Конструкторы

        public OS() { }

        /// <summary>
        /// Конструктор для заполнения базы данных.
        /// </summary>
        /// <param name="inv"></param>
        /// <param name="name"></param>
        /// <param name="type"></param>
        /// <param name="marka"></param>
        /// <param name="SN"></param>
        /// <param name="Date"></param>
        /// <param name="status"></param>
        /// <param name="otvetstvenniy"></param>
        /// <param name="podrazdelenie"></param>
        /// <param name="otdel"></param>
        /// <param name="statusRecList"></param>
        /// <param name="statusrec"></param>
        public OS(string inv, string name, string type, string marka, string SN,
            string Date, string status, string user, string otvetstvenniy, string podrazdelenie, string otdel, List<string> statusRecList, string statusrec)
        {
            InvNum = inv;
            Name = name;
            TypeOs = type;
            Marka = marka;
            SerialNum = SN;
            DatePostanovki = Date;
            Sostoyanie = status;
            User = user;
            Otvetstvenniy = otvetstvenniy;
            Podrazdelenie = podrazdelenie;
            Otdel = otdel;
            StatusRecList = statusRecList;
            StatusRec = statusrec;
        }

        /// <summary>
        /// Конструктор для создания объектов из Excel файла.
        /// </summary>
        /// <param name="rownum"></param>
        /// <param name="inv"></param>
        /// <param name="name"></param>
        /// <param name="type"></param>
        /// <param name="marka"></param>
        /// <param name="SN"></param>
        /// <param name="Date"></param>
        /// <param name="status"></param>
        /// <param name="otvetstvenniy"></param>
        /// <param name="podrazdelenie"></param>
        /// <param name="otdel"></param>
        /// <param name="statusRecList"></param>
        /// <param name="repeat"></param>
        /// <param name="comment"></param>
        /// <param name="loc"></param>
        /// <param name="user"></param>
        /// <param name="newotv"></param>
        /// <param name="factsn"></param>
        public OS(int rownum, string inv, string name, string type, string marka, string SN,
            string Date, string status, string user, string otvetstvenniy, string podrazdelenie, string otdel,
            List<string> statusRecList, string statusrec, string repeat, string comment, string loc, string newotv, string factsn)
            : this(inv, name, type, marka, SN, Date, status, user, otvetstvenniy, podrazdelenie, otdel, statusRecList, statusrec)
        {
            NumRow = rownum;
            RepeatRec = repeat;
            Comment = comment;
            Location = loc;
            NewOtv = newotv;
            FactSerialNum = factsn;
        }

        /// <summary>
        /// Конструктор всех данных, кроме порядкового номера. Для заполнения следующим конструктором из существующего объекта.
        /// </summary>
        /// <param name="inv"></param>
        /// <param name="name"></param>
        /// <param name="type"></param>
        /// <param name="marka"></param>
        /// <param name="SN"></param>
        /// <param name="Date"></param>
        /// <param name="status"></param>
        /// <param name="otvetstvenniy"></param>
        /// <param name="podrazdelenie"></param>
        /// <param name="otdel"></param>
        /// <param name="statusRec"></param>
        /// <param name="repeat"></param>
        /// <param name="comment"></param>
        /// <param name="loc"></param>
        /// <param name="user"></param>
        /// <param name="newotv"></param>
        /// <param name="factsn"></param>
        /// <param name="statusrec"></param>
        /// <param name="colorcell"></param>
        public OS(string inv, string name, string type, string marka, string SN,
            string Date, string status, string otvetstvenniy, string podrazdelenie, string otdel,
            List<string> statusRec, string repeat, string comment, string loc, string user, string newotv, string factsn, string statusrec)
        {
            InvNum = inv;
            Name = name;
            TypeOs = type;
            Marka = marka;
            SerialNum = SN;
            DatePostanovki = Date;
            Sostoyanie = status;
            Otvetstvenniy = otvetstvenniy;
            Podrazdelenie = podrazdelenie;
            Otdel = otdel;
            StatusRecList = statusRec;
            RepeatRec = repeat;
            Comment = comment;
            Location = loc;
            User = user;
            NewOtv = newotv;
            FactSerialNum = factsn;
            StatusRec = statusrec;
        }

        /// <summary>
        /// Конструктор создания экземпляра из другого экземпляра без порядкового номера.
        /// </summary>
        /// <param name="obj"></param>
        public OS(OS obj) : this(obj.InvNum, obj.Name, obj.TypeOs, obj.Marka, obj.SerialNum,
            obj.DatePostanovki, obj.Sostoyanie, obj.Otvetstvenniy, obj.Podrazdelenie, obj.Otdel,
            obj.StatusRecList, obj.RepeatRec, obj.Comment, obj.Location, obj.User, obj.NewOtv, obj.FactSerialNum, obj.StatusRec)
        {

        }

        #endregion

        #region Методы

        /// <summary>
        /// Заполняет коллекцию List данными из Excel файла. Возвращает значение, указывающее на успешность выполнения.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="database"></param>
        /// <returns></returns>
        public static bool CreateDataBase(string filepath, List<OS> database)
        {
            //try
            //{
                // Инициализация потока чтения файла.
                using (FileStream stream = File.Open(filepath, FileMode.Open, FileAccess.Read))
                {
                    // Инициализация считывателя данных Excel.
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                var item = new OS(
                                    Convert.ToString(reader.GetValue(0)),
                                    Convert.ToString(reader.GetValue(1)),
                                    Convert.ToString(reader.GetValue(2)),
                                    Convert.ToString(reader.GetValue(3)),
                                    Convert.ToString(reader.GetValue(4)),
                                    Convert.ToString(reader.GetValue(5)),
                                    Convert.ToString(reader.GetValue(6)),
                                    Convert.ToString(reader.GetValue(8)),
                                    Convert.ToString(reader.GetValue(7)),
                                    Convert.ToString(reader.GetValue(9)),
                                    Convert.ToString(reader.GetValue(10)),
                                    statusOs,
                                    string.Empty
                                    );
                                database.Add(item);
                            }
                        } while (reader.NextResult());
                    }
                }
                return true;
            //}
            //catch
            //{
            //    //Вернуть false.
            //    return false;
            //}
        }

        /// <summary>
        /// Десериализует XML файл в коллекцию ObservableCollection<OS> с OpenFileDialog. Возвращает значение, указывающее на успешность выполнения.
        /// </summary>
        /// <returns></returns>
        public static bool OpenXmlWithFileDialogASync()
        {
            // Диалоговое окно открытия файла.
            var ofd = new OpenFileDialog
            {
                // Настройка файл диалога.
                DefaultExt = "*.xml",
                Filter = "XML файлы (*.xml)|*.xml",
                Title = "Открыть файл инвентаризации"
            };
            // Если результат диалога положителен.
            if (ofd.ShowDialog() == true)
            {
                return OpenXml(ofd.FileName);
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Десериализует XML файл в коллекцию ObservableCollection<OS>. Возвращает значение, указывающее на успешность выполнения.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool OpenXml(string filePath)
        {
            try
            {
                // Передаем в конструктор тип класса.
                var opener = new XmlSerializer(typeof(ObservableCollection<OS>));

                // Десериализация.
                using (var fs = new FileStream(filePath, FileMode.Open))
                {
                    ScanList.Clear();
                    ScanList = (ObservableCollection<OS>)opener.Deserialize(fs);
                    ForWorks.viewSource.Source = ScanList;
                    // Подписка на события изменения базы данных ObservableCollection.
                    ScanList.CollectionChanged += MainWindow.SelfRef.ChangeCollection;
                    MainWindow.SelfRef.SetFilters();
                }
                // Установить открытый файл как текущий.
                ForWorks.currentWorkFile = filePath;
                // Установка переменной состояния изменений документа в ложь.
                ForWorks.checkEdit = false;
                // Установить счетчик строк в зависимости от наличия элементов в читаемом файле.
                RowNumCounter = ScanList.Count > 0 ? ScanList.Last().NumRow + 1 : 1;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Сериализует коллекцию отсканированных элементов в XML файл с SaveFileDialog. Возвращает значение, указывающее на успешность выполнения.
        /// </summary>
        /// <returns></returns>
        public static bool SaveXmlWithFileDialog()
        {
            // Открыть диалог сохранения файла.
            SaveFileDialog sfd = new SaveFileDialog
            {
                // Параметры диалога.
                DefaultExt = "*.xml",
                Filter = "XML файлы |*.xml",
                Title = "Сохранение данных инвентаризации"
            };
            if (sfd.ShowDialog() == true)
            {
                return SaveXml(sfd.FileName);
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Сериализует коллекцию отсканированных элементов в XML файл. Возвращает значение, указывающее на успешность выполнения.
        /// </summary>
        /// <param name="filePath"></param>
        public static bool SaveXml(string filePath)
        {
            try
            {
                // передаем в конструктор тип класса
                var saver = new XmlSerializer(typeof(ObservableCollection<OS>));

                // получаем поток, куда будем записывать сериализованный объект
                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    saver.Serialize(fs, ScanList);
                }
                // Установить сохраненный файл как текущий.
                ForWorks.currentWorkFile = filePath;
                // Установка переменной состояния изменений документа в ложь.
                ForWorks.checkEdit = false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Сериализует текущую коллекцию отсканированных элементов в XML файл по таймеру.
        /// </summary>
        /// <param name="obj"></param>
        public static void SaveTempXml(object obj)
        {
            try
            {
                var filepath = (string)obj;
                // передаем в конструктор тип класса
                var saver = new XmlSerializer(typeof(ObservableCollection<OS>));

                // получаем поток, куда будем записывать сериализованный объект
                using (var fs = new FileStream(filepath, FileMode.Create))
                {
                    saver.Serialize(fs, ScanList);
                }
                // Установить сохраненный файл как текущий.
                ForWorks.currentWorkFile = filepath;
                // Установка переменной состояния изменений документа в ложь.
                ForWorks.checkEdit = false;
            }
            catch
            {
                MessageBox.Show("Ошибка записи текущего файла!\nПерезапишите текущий файл!", "Сохранение текущего файла", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Сохраняет отображаемые инвентарные номера в текстовый файл.
        /// </summary>
        public static void SaveInvNumToTxt()
        {
            // Инициализация окна сохранения.
            SaveFileDialog saveTxtdialog = new SaveFileDialog
            {
                // Параметры диалога.
                DefaultExt = "*.txt",
                Filter = "Текстовые файлы |*.txt",
                Title = "Сохранение инвентарных номеров"
            };
            // Если нажата кнопка сохранить.
            if (saveTxtdialog.ShowDialog() == true)
            {
                // Используем поток для записи.
                using (StreamWriter savetxt = new StreamWriter(saveTxtdialog.FileName, false, System.Text.Encoding.Default))
                {
                    try
                    {
                        // Перебрать элементы текущего отображения.
                        foreach (object o in ForWorks.viewSource.View)
                        {
                            if (o is OS obj)
                            {
                                // Записать в строку инвентарный номер.
                                savetxt.WriteLine(obj.InvNum);
                            }
                        }
                        // Вывести сообщение по окончанию записи.
                        MessageBox.Show("Инвентарные номера сохранены", "Сохранения инвентарный номеров", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch
                    {
                        // Вывести сообщение об ошибке.
                        MessageBox.Show("Ошибка сохранения", "Сохранения инвентарный номеров", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Сохраняет данные сканирования в Excel файл. Возвращает результат выполнения.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="colHeaders"></param>
        /// <param name="view"></param>
        /// <returns></returns>
        public static bool SaveToExcel(string filePath, List<string> colHeaders, ICollectionView view)
        {
            try
            {
                // Новый экхемпляр книги.
                using (XLWorkbook wb = new XLWorkbook())
                {
                    // Новая вкладка.
                    IXLWorksheet ws = wb.Worksheets.Add("Отчет");
                    // Переменная количества требуемых столбцов в файле.
                    int columnCount = 19;
                    // Перебираем столбцы датагрида.
                    for (int i = 1; i < columnCount; i++)
                    {
                        // В первой строке записываем названия столбцов.
                        ws.Cell(1, i).Value = colHeaders[i - 1];
                    }
                    // Добавляем столбец для обозначения повторных записей.
                    ws.Cell(1, 19).Value = "Повторная запись";
                    // Счетчик строк, начиная со второй.
                    int rowNum = 2;
                    // Перебираем текущее представление датагрида.
                    foreach (var temp in view)
                    {
                        // Если элемент представления можно привести к типу OS.
                        if (temp is OS item)
                        {
                            // Цикл счетчик по количеству столбцов.
                            for (int i = 1; i <= columnCount; i++)
                            {
                                ws.Cell(rowNum, 1).Value = item.NumRow;
                                ws.Cell(rowNum, 2).Style.NumberFormat.Format = "#";
                                ws.Cell(rowNum, 2).Value = item.InvNum;
                                ws.Cell(rowNum, 3).Value = item.Name;
                                ws.Cell(rowNum, 4).Value = item.TypeOs;
                                ws.Cell(rowNum, 5).Value = item.Marka;
                                ws.Cell(rowNum, 6).Value = item.SerialNum;
                                ws.Cell(rowNum, 7).Value = item.DatePostanovki;
                                ws.Cell(rowNum, 8).Value = item.Sostoyanie;
                                ws.Cell(rowNum, 9).Value = item.Otvetstvenniy;
                                ws.Cell(rowNum, 10).Value = item.Podrazdelenie;
                                ws.Cell(rowNum, 11).Value = item.Otdel;
                                ws.Cell(rowNum, 12).Value = item.Comment;
                                ws.Cell(rowNum, 13).Value = item.Location;
                                ws.Cell(rowNum, 14).Value = item.User;
                                ws.Cell(rowNum, 15).Value = item.NewUser;
                                ws.Cell(rowNum, 16).Value = item.NewOtv;
                                ws.Cell(rowNum, 17).Value = item.FactSerialNum;
                                ws.Cell(rowNum, 18).Value = item.StatusRec;
                                ws.Cell(rowNum, 19).Value = item.RepeatRec;
                            }
                            // Увеличить счетчик строк.
                            rowNum++;
                        }
                    }

                    // Стиль для заголовков и задание его параметров.
                    IXLStyle titlesStyle = wb.Style;
                    titlesStyle.Font.Bold = true;
                    titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titlesStyle.Fill.BackgroundColor = XLColor.BlueBell;
                    titlesStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    titlesStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    titlesStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    titlesStyle.Border.BottomBorder = XLBorderStyleValues.Thin;

                    int rowscount = ws.Rows().Count();

                    // Применение стилей к диапазонам.
                    ws.Range(1, 1, 1, columnCount).Style = titlesStyle;
                    ws.Range(2, 1, rowscount, columnCount).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    ws.Range(2, 1, rowscount, columnCount).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    ws.Range(2, 1, rowscount, columnCount).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    ws.Range(2, 1, rowscount, columnCount).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    ws.Range(2, 1, rowscount, columnCount).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    ws.Range(2, 1, rowscount, columnCount).SetAutoFilter();

                    for (int i = 1; i < ws.Rows().Count(); i++)
                    {
                        if (i % 2 == 0)
                        {
                            ws.Row(i).Style.Fill.BackgroundColor = XLColor.FromHtml("#ADD8E6");
                        }
                    }
                    // Автоподбор ширины столбцов.
                    ws.Columns().AdjustToContents();
                    // Сохранения файла.
                    wb.SaveAs(filePath);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Возвращает свойство объекта по имени.
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public PropertyInfo GetProperty(string propertyName)
        {
            return typeof(OS).GetProperty(propertyName);
        }

        #endregion

        #region События

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        #endregion
    }
}