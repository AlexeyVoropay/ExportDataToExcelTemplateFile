using System;

namespace ExcelExport.Models
{
    /// <summary>
    /// Значение для вставки в ячейку
    /// </summary>
    public class ValueToInsert
    {
        /// <summary>
        /// Наименование значения в ячейке, в которое будет вставлено значение
        /// </summary>
        public string FieldName { get; set; }
        /// <summary>
        /// Тип значения
        /// </summary>
        public Type Type { get; set; }

        /// <summary>
        /// Значение
        /// </summary>
        public object Value { get; set; }
        /// <summary>
        /// Формула (без знака "=")
        /// </summary>
        public bool IsFormula { get; set; }
        /// <summary>
        /// Ссылка на ячейку для копирования стиля
        /// </summary>
        public string CellReferenceStyle { get; set; }

        public ValueToInsert()
        { }

        public ValueToInsert(string fieldName, Type fieldType,  object fieldValue, bool isFormula = false, string cellReferenceStyle = null)
        {
            FieldName = fieldName;
            Type = fieldType;
            Value = fieldValue;
            IsFormula = isFormula;
            CellReferenceStyle = cellReferenceStyle;
        }
    }
}