using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;

namespace DeepDuplicateFinder
{
    public class UnifiedReportGenerator
    {
        public string CreateReport(
            ObjectInfo folderInfo,
            int totalObjects,
            int totalDetails,
            int totalPreparations,
            int totalMaterials,
            int totalProblemObjects,
            List<ReportItem> reportItems)
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string safeName = ReplaceInvalidChars(folderInfo.Name);
                string fileName = $"materials_report_{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.html";
                string filePath = Path.Combine(desktop, fileName);

                var sb = new StringBuilder();

                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset='utf-8'>");
                sb.AppendLine("<title>Отчет о множественных материалах</title>");
                sb.AppendLine("<style>");
                sb.AppendLine("body { font-family: Arial; margin: 20px; background: #f5f5f5; }");
                sb.AppendLine(".container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; }");
                sb.AppendLine("h1 { color: #333; text-align: center; }");
                sb.AppendLine(".summary { background: #f0f0f0; padding: 15px; border-radius: 5px; margin-bottom: 20px; }");
                sb.AppendLine(".stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px,1fr)); gap: 10px; margin-bottom: 20px; }");
                sb.AppendLine(".stat-card { background: white; padding: 15px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); text-align: center; }");
                sb.AppendLine(".stat-number { font-size: 24px; font-weight: bold; color: #2196F3; }");
                sb.AppendLine(".detail-det { border-left: 4px solid #4CAF50; background: #f9f9f9; margin: 15px 0; padding: 15px; border-radius: 5px; }");
                sb.AppendLine(".detail-prep { border-left: 4px solid #2196F3; background: #f9f9f9; margin: 15px 0; padding: 15px; border-radius: 5px; }");
                sb.AppendLine(".material-group { background: #fff3e0; margin: 10px 0; padding: 10px; border-radius: 5px; }");
                sb.AppendLine("table { width: 100%; border-collapse: collapse; margin-top: 10px; }");
                sb.AppendLine("th { background: #4CAF50; color: white; padding: 10px; text-align: left; }");
                sb.AppendLine("td { padding: 8px; border-bottom: 1px solid #ddd; }");
                sb.AppendLine("tr:hover { background: #f5f5f5; }");
                sb.AppendLine("</style>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");
                sb.AppendLine("<div class='container'>");

                // Заголовок
                sb.AppendLine($"<h1>Отчет о множественных материалах</h1>");
                sb.AppendLine($"<p><strong>Папка:</strong> {EscapeHtml(folderInfo.Name)} (ID: {folderInfo.Id})</p>");
                sb.AppendLine($"<p><strong>Дата:</strong> {DateTime.Now:dd.MM.yyyy HH:mm:ss}</p>");

                // Сводка
                sb.AppendLine("<div class='summary'>");
                sb.AppendLine("<h2>Сводка</h2>");
                sb.AppendLine("<div class='stats'>");
                sb.AppendLine($"<div class='stat-card'><div class='stat-number'>{totalObjects}</div><div>Всего объектов</div></div>");
                sb.AppendLine($"<div class='stat-card'><div class='stat-number'>{totalDetails}</div><div>Деталей</div></div>");
                sb.AppendLine($"<div class='stat-card'><div class='stat-number'>{totalPreparations}</div><div>Заготовок</div></div>");
                sb.AppendLine($"<div class='stat-card'><div class='stat-number'>{totalMaterials}</div><div>Связей с материалами</div></div>");
                sb.AppendLine($"<div class='stat-card'><div class='stat-number'>{totalProblemObjects}</div><div>Объектов с >1 материалом</div></div>");
                sb.AppendLine("</div>");
                sb.AppendLine("</div>");

                if (reportItems.Count == 0)
                {
                    sb.AppendLine("<div style='text-align: center; padding: 40px; background: #e8f5e8; border-radius: 5px;'>");
                    sb.AppendLine("<h2 style='color: #4CAF50;'>✅ Объекты с множественными материалами не найдены</h2>");
                    sb.AppendLine("</div>");
                }
                else
                {
                    sb.AppendLine($"<h2>Найдено объектов с множественными материалами: {reportItems.Count}</h2>");

                    foreach (var item in reportItems)
                    {
                        string detailClass = item.ObjectTypeName == "Деталь" ? "detail-det" : "detail-prep";

                        sb.AppendLine($"<div class='{detailClass}'>");
                        sb.AppendLine($"<h3>{item.ObjectTypeName}: {EscapeHtml(item.Object.Name)} (ID: {item.Object.Id})</h3>");
                        sb.AppendLine($"<p><strong>Версия:</strong> {EscapeHtml(item.Object.Version)}</p>");
                        sb.AppendLine($"<p><strong>Тип материала:</strong> {item.MaterialTypeName}</p>");
                        sb.AppendLine($"<p><strong>Всего материалов:</strong> {item.TotalMaterials}</p>");

                        sb.AppendLine("<h4>Все материалы:</h4>");
                        sb.AppendLine("<table>");
                        sb.AppendLine("<tr><th>ID</th><th>Название</th><th>Версия</th><th>Тип</th><th>LinkId</th></tr>");

                        foreach (var m in item.Materials)
                        {
                            sb.AppendLine("<tr>");
                            sb.AppendLine($"<td>{m.Id}</td>");
                            sb.AppendLine($"<td>{EscapeHtml(m.Name)}</td>");
                            sb.AppendLine($"<td>{EscapeHtml(m.Version)}</td>");
                            sb.AppendLine($"<td>{EscapeHtml(m.Type)}</td>");
                            sb.AppendLine($"<td>{m.LinkId}</td>");
                            sb.AppendLine("</tr>");
                        }
                        sb.AppendLine("</table>");

                        if (item.MaterialGroups.Count > 0)
                        {
                            sb.AppendLine("<h4>Дублирующиеся материалы:</h4>");
                            foreach (var g in item.MaterialGroups)
                            {
                                sb.AppendLine("<div class='material-group'>");
                                sb.AppendLine($"<p><strong>{EscapeHtml(g.MaterialName)}</strong> (Версия: {EscapeHtml(g.MaterialVersion)}) → {g.LinkCount} раз(а)</p>");
                                sb.AppendLine("</div>");
                            }
                        }

                        sb.AppendLine("</div>");
                    }
                }

                sb.AppendLine("</div>");
                sb.AppendLine("</body>");
                sb.AppendLine("</html>");

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                return filePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка создания отчета: {ex.Message}");
            }
        }

        private string EscapeHtml(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";
            return System.Net.WebUtility.HtmlEncode(text);
        }

        private string ReplaceInvalidChars(string filename, string replacement = "_")
        {
            if (string.IsNullOrEmpty(filename)) return "unknown";
            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                filename = filename.Replace(c.ToString(), replacement);
            }
            if (filename.Length > 50)
                filename = filename.Substring(0, 50);
            return filename.Trim();
        }
    }
}