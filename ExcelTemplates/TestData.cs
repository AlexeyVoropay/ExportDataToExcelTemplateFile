using ExcelTemplates.Models;
using System.Collections.Generic;

namespace ExcelTemplates
{
    public static class TestData
    {
        public static DrillingReport GetTestData()
        {
            return new DrillingReport
            {
                ReportDate = "12.07.2020",
                ReportNumber = "87",                
                WellInfo = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("Заказчик","ООО\"ИНК\""),
                    new KeyValuePair<string, string>("Месторождение","Ярактинское"),
                    new KeyValuePair<string, string>("Куст","16"),
                    new KeyValuePair<string, string>("№ скважины","9636"),
                },
                SvInfo = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("Супервайзер","Шарафутдинов Радик Мансурович"),
                    new KeyValuePair<string, string>("Телефон","+7(963)-450-10-87"),
                    new KeyValuePair<string, string>("e-mail","Sharafutdinov.RM@mail.ru"),
                },
                Сonstruction = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("Направление",""),
                    new KeyValuePair<string, string>("Кондуктор",""),
                    new KeyValuePair<string, string>("Техническая колонна",""),
                    new KeyValuePair<string, string>("Эксплуатационная колонна",""),
                    new KeyValuePair<string, string>("Хвостовик",""),
                    new KeyValuePair<string, string>("Пилотный ствол",""),
                },
                Hse = new Hse
                {
                    NumStopCards = 12,
                    NumAlarmsDone = 3,
                    LastSafetyMeeting = "18.08.2020",
                    Incident = "Что-то случилось",
                },
            };
        }
    }
}