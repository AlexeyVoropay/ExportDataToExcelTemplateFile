namespace ExcelExport.Models
{
    public class LocationWithValue
    {
        /// <summary>
        /// Индекс строки
        /// </summary>
        public uint RowIndex { get; private set; }

        /// <summary>
        /// Индекс колонки
        /// </summary>
        public string ColumnIndex { get; private set; }

        /// <summary>
        /// Наименование переменной
        /// </summary>
        public string ValueName { get; private set; }

        public LocationWithValue(uint row, string column, string field)
        {
            RowIndex = row;
            ColumnIndex = column;
            ValueName = field;
        }
    }
}
