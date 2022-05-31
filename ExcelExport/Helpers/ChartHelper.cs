using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ExcelExport.Helpers
{
    public static class ChartHelper
    {
        public static ChartPart GetChartPartByTitle(IEnumerable<ChartPart> chartParts, string title)
        {
            foreach (var item in chartParts)
            {
                var chartSpace = item.ChartSpace;
                var chart = chartSpace.GetFirstChild<Chart>();
                var chartTitle = chart.Title;
                if (chartTitle != null)
                {
                    if (chartTitle.InnerText.StartsWith(title))
                    {
                        return item;
                    }
                }
                else if (chart.InnerText.StartsWith(title))
                {
                    return item;
                }
            }
            return null;
        }

        public static ScatterChartSeries GetScatterChartSeriesBySeriesText(IEnumerable<ScatterChartSeries> items, string seriesText)
        {
            foreach (var item in items)
            {
                var firstSeriesText = item.GetFirstChild<SeriesText>();
                if (firstSeriesText.InnerText == seriesText)
                {
                    return item;
                }
            }
            return null;
        }

        public static string ChangeFormula(string formula, int pointSkip, int pointCount)
        {
            if (pointCount > 0)
            {
                var s = formula.Split('$');
                var sheet = s[0];
                var firstCell = $"{s[1]}${s[2]}".Replace(":", "");
                var newFirstCell = $"{s[1]}${int.Parse(s[2].Replace(":", "")) + pointSkip}";
                var lastCell = $"{s[1]}${int.Parse(s[2].Replace(":", "")) + pointCount - 1}";
                formula = $"{sheet}${newFirstCell}:{lastCell}";
            }
            return formula;
        }
    }
}