using ExcelTemplates.TemplatesModels;
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
                Knbk = new List<KnbkItem>
                {                    
                    new KnbkItem{ Name = "Долото PDC XS613", In = "-"     , Od = "155,6"    , Connection = "Н88 Нипель"              ,                   Len = "0,2     ",TotalLen="0,2"   },
                    new KnbkItem{ Name = "ДРУ2-120РСФКТМ",  In = "51"    , Od = "120"      , Connection = "Н102 Нипель / М 88 Муфта",                   Len = "8,62    ",TotalLen="8,82"  },
                    new KnbkItem{ Name = "КС-150",          In = "67"    , Od = "150"      , Connection = "Н102 Нипель / М102 Муфта",                   Len = "0,27    ",TotalLen="9,09"  },
                    new KnbkItem{ Name = "Переводник UBHO", In = "70"    , Od = "121"      , Connection = "Н102 Нипель / М102 Муфта",                   Len = "0,98    ",TotalLen="10,07" },
                    new KnbkItem{ Name = "НУБТ (телесистема)", In = "70"    , Od = "120"      , Connection = "Н102 Нипель / М102 Муфта",                   Len = "7,27    ",TotalLen="17,34" },
                    new KnbkItem{ Name = "НУБТ короткая", In = "51"    , Od = "120"      , Connection = "Н102 Нипель / М102 Муфта",                   Len = "4,75    ",TotalLen="22,09" },
                    new KnbkItem{ Name = "КС-150", In = "57"    , Od = "150"      , Connection = "Н102 Нипель / М102 Муфта",                   Len = "0,27    ",TotalLen="22,36" },
                    new KnbkItem{ Name = "НУБТ", In = "57,2"  , Od = "120"      , Connection = "Н102 Нипель / М102 Муфта",                   Len = "8,8     ",TotalLen="31,16" },
                    new KnbkItem{ Name = "ТБТ-89 (2 тр.)       ", In = "64"    , Od = "89"       , Connection = "Н102 Нипель / М102 Муфта",                   Len = "18,46   ",TotalLen="49,62" },
                    new KnbkItem{ Name = "СБТ-89(14тр.)", In = "57,2"  , Od = "89"       , Connection = "Н102 Нипель / М102 Муфта",                   Len = "133,16  ",TotalLen="182,78"},
                },
                Trajectory = new List<TrajectoryItem>
                {
                    new TrajectoryItem{ Md = "2500", Incl = "76,5", Azi = "213,5",Tvd ="2000",Closure ="500", Dls = "1,5", Compare = "0,5м выше / 0,5м правее"},
                },
                GtiSummaryDuration = "24",
                Gti = new List<GtiItem>
                {
                    new GtiItem{ StartTime = "0:00", EndTime = "1:00", Duration = "1:00", Duration2 = "1", StartDepth = "2488", EndDepth = "2500", 
                        Operation = "Механическое бурение", Modes = "Q=36лс, G=5т, P=150атм, N=80об/мин, М=15кH*м"},
                    new GtiItem{ StartTime = "1:00", EndTime = "1:15", Duration = "0:15", Duration2 = "0,25", StartDepth = "2500", EndDepth = "2500",
                        Operation = "Проработка перед наращиванием"},
                    new GtiItem{ StartTime = "1:15", EndTime = "2:00", Duration = "0:45", Duration2 = "0,75", StartDepth = "2500", EndDepth = "2500",
                        Operation = "Промежуточная промывка в открытом стволе"},
                    new GtiItem{ StartTime = "2:00", EndTime = "0:00", Duration = "22:00", Duration2 = "22", StartDepth = "2500", EndDepth = "2500",
                        Operation = "Ремонт бурового насоса",
                        NptCategory = "Ремонт оборудования", NptDuration = "22", NptResponsible="ООО \"ИНК-Сервис\"", Comment = "Промыло гидравлику"},
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