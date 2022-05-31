namespace ExcelExport.Models
{
    public class ChartToInsert
    {
        public string ChartTitle { get; set; }
        public int PointCount { get; set; }
        public int MoveOnRows { get; set; }
        /// <summary>
        /// Признак что график надо скрыть
        /// </summary>
        public bool IsHide { get; set; }
        public int StartRowIndex { get; set; }
        public int EndRowIndex { get; set; }
        /// <summary>
        /// Цвета для столбцов
        /// </summary>
        public string[] DataPointsColors { get; set; }
    }
}