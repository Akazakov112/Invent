using System.Collections.ObjectModel;

namespace Invent
{
    /// <summary>
    /// Класс фильтров.
    /// </summary>
    class Filters
    {
        /// <summary>
        /// Коллекция значений свойства "Ответственный" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterOtv { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Новый ответственный" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterNewOtv { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Пользователь" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterUser { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Местоположение" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterLocation { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Подразделение" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterPodrazdelenie { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Тип ОС" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterTypeOs { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Состояние" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterSostoyanie { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений вариантов "Тип инвентарного номера" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterTypeInvNumn { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Дата постановки" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterDate { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Отдел местоположения" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterOtdel { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Коллекция значений свойства "Статус записи" у объектов класса OS для фильтра.
        /// </summary>
        public static ObservableCollection<string> FilterStatusRec { get; set; } = new ObservableCollection<string>();


        /// <summary>
        /// Заполняет фильтры уникальными значениями свойств.
        /// </summary>
        public static void GetFilters()
        {
            // Очистить все фильтры.
            FilterOtv.Clear();
            FilterNewOtv.Clear();
            FilterUser.Clear();
            FilterLocation.Clear();
            FilterPodrazdelenie.Clear();
            FilterTypeOs.Clear();
            FilterSostoyanie.Clear();
            FilterDate.Clear();
            FilterOtdel.Clear();
            FilterTypeInvNumn.Clear();
            FilterStatusRec.Clear();

            // Добавить во все фильтры пункты Все и Пустые значения.
            FilterOtv.Add("Все");
            FilterOtv.Add("Пустые значения");
            FilterNewOtv.Add("Все");
            FilterNewOtv.Add("Пустые значения");
            FilterUser.Add("Все");
            FilterUser.Add("Пустые значения");
            FilterLocation.Add("Все");
            FilterLocation.Add("Пустые значения");
            FilterPodrazdelenie.Add("Все");
            FilterPodrazdelenie.Add("Пустые значения");
            FilterTypeOs.Add("Все");
            FilterTypeOs.Add("Пустые значения");
            FilterSostoyanie.Add("Все");
            FilterSostoyanie.Add("Пустые значения");
            FilterDate.Add("Все");
            FilterDate.Add("Пустые значения");
            FilterOtdel.Add("Все");
            FilterOtdel.Add("Пустые значения");
            FilterStatusRec.Add("Все");
            FilterStatusRec.Add("Пустые значения");
            FilterTypeInvNumn.Add("Все");

            foreach (object o in ForWorks.viewSource.View)
            {
                if (o is OS obj)
                {
                    if (!FilterOtv.Contains(obj.Otvetstvenniy) && !string.IsNullOrWhiteSpace(obj.Otvetstvenniy))
                    {
                        FilterOtv.Add(obj.Otvetstvenniy);
                    }

                    if (!FilterNewOtv.Contains(obj.NewOtv) && !string.IsNullOrWhiteSpace(obj.NewOtv))
                    {
                        FilterNewOtv.Add(obj.NewOtv);
                    }

                    if (!FilterUser.Contains(obj.User) && !string.IsNullOrWhiteSpace(obj.User))
                    {
                        FilterUser.Add(obj.User);
                    }

                    if (!FilterLocation.Contains(obj.Location) && !string.IsNullOrWhiteSpace(obj.Location))
                    {
                        FilterLocation.Add(obj.Location);
                    }

                    if (!FilterPodrazdelenie.Contains(obj.Podrazdelenie) && !string.IsNullOrWhiteSpace(obj.Podrazdelenie))
                    {
                        FilterPodrazdelenie.Add(obj.Podrazdelenie);
                    }

                    if (!FilterTypeOs.Contains(obj.TypeOs) && !string.IsNullOrWhiteSpace(obj.TypeOs))
                    {
                        FilterTypeOs.Add(obj.TypeOs);
                    }

                    if (!FilterSostoyanie.Contains(obj.Sostoyanie) && !string.IsNullOrWhiteSpace(obj.Sostoyanie))
                    {
                        FilterSostoyanie.Add(obj.Sostoyanie);
                    }

                    if (!FilterDate.Contains(obj.DatePostanovki) && !string.IsNullOrWhiteSpace(obj.DatePostanovki))
                    {
                        FilterDate.Add(obj.DatePostanovki);
                    }

                    if (!FilterOtdel.Contains(obj.Otdel) && !string.IsNullOrWhiteSpace(obj.Otdel))
                    {
                        FilterOtdel.Add(obj.Otdel);
                    }

                    if (!FilterStatusRec.Contains(obj.StatusRec) && !string.IsNullOrWhiteSpace(obj.StatusRec))
                    {
                        FilterStatusRec.Add(obj.StatusRec);
                    }

                    if (obj.InvNum.StartsWith("29600") && !FilterTypeInvNumn.Contains("Деловые линии"))
                    {
                        FilterTypeInvNumn.Add("Деловые линии");
                    }

                    if (obj.InvNum.StartsWith("296CC") && !FilterTypeInvNumn.Contains("КЦ"))
                    {
                        FilterTypeInvNumn.Add("КЦ");
                    }

                    if (obj.InvNum.StartsWith("6000") && !FilterTypeInvNumn.Contains("ДЛ Транс"))
                    {
                        FilterTypeInvNumn.Add("ДЛ Транс");
                    }

                    if (obj.InvNum.StartsWith("11100") && !FilterTypeInvNumn.Contains("Гет Карго"))
                    {
                        FilterTypeInvNumn.Add("Гет Карго");
                    }
                }
            }
        }
    }
}