using HtmlAgilityPack;
using System.Data;
using System.Linq;

namespace HoiThao_Core.Helpers
{
    public static class HtmlTableToDataTable
    {
        public static DataTable ConvertToDataTable(string htmlCode)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(htmlCode);
            var headers = doc.DocumentNode.SelectNodes("//tr/th");
            DataTable table = new DataTable();
            foreach (HtmlNode header in headers)
                table.Columns.Add(header.InnerText); // create columns from th
                                                     // select rows with td elements
            foreach (var row in doc.DocumentNode.SelectNodes("//tr[td]"))
                table.Rows.Add(row.SelectNodes("td").Select(td => td.InnerText).ToArray());

            return table;
        }
    }
}