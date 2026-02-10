using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
namespace DeepDuplicateFinder
{
    public class HtmlReportGenerator
    {
        public string CreateHtmlReport(ObjectInfo folderInfo,
                                     int totalObjectsCount,
                                     int objectsWithNamesCount,
                                     List<IGrouping<string, ObjectInfo>> fullDuplicates,
                                     List<IGrouping<string, ObjectInfo>> nameTypeDuplicates,
                                     int totalRealDuplicates)
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string safeName = ReplaceInvalidChars(folderInfo.Name);
                string fileName = $"duplicates_report_{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.html";
                string filePath = Path.Combine(desktop, fileName);

                var sb = new StringBuilder();

                // Начало HTML документа
                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("<html lang='ru'>");
                sb.AppendLine("<head>");
                sb.AppendLine("    <meta charset='UTF-8'>");
                sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
                sb.AppendLine($"    <title>Отчет о дубликатах - {EscapeHtml(folderInfo.Name)}</title>");
                sb.AppendLine("    <style>");
                sb.AppendLine("        * {");
                sb.AppendLine("            margin: 0;");
                sb.AppendLine("            padding: 0;");
                sb.AppendLine("            box-sizing: border-box;");
                sb.AppendLine("            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;");
                sb.AppendLine("        }");
                sb.AppendLine("        body {");
                sb.AppendLine("            background-color: #f5f5f5;");
                sb.AppendLine("            color: #333;");
                sb.AppendLine("            line-height: 1.6;");
                sb.AppendLine("            padding: 20px;");
                sb.AppendLine("        }");
                sb.AppendLine("        .container {");
                sb.AppendLine("            max-width: 1200px;");
                sb.AppendLine("            margin: 0 auto;");
                sb.AppendLine("            background-color: white;");
                sb.AppendLine("            border-radius: 10px;");
                sb.AppendLine("            box-shadow: 0 2px 15px rgba(0,0,0,0.1);");
                sb.AppendLine("            padding: 30px;");
                sb.AppendLine("        }");
                sb.AppendLine("        .header {");
                sb.AppendLine("            text-align: center;");
                sb.AppendLine("            margin-bottom: 30px;");
                sb.AppendLine("            padding-bottom: 20px;");
                sb.AppendLine("            border-bottom: 3px solid #4CAF50;");
                sb.AppendLine("        }");
                sb.AppendLine("        .header h1 {");
                sb.AppendLine("            color: #2c3e50;");
                sb.AppendLine("            font-size: 28px;");
                sb.AppendLine("            margin-bottom: 10px;");
                sb.AppendLine("        }");
                sb.AppendLine("        .header .date {");
                sb.AppendLine("            color: #7f8c8d;");
                sb.AppendLine("            font-size: 14px;");
                sb.AppendLine("        }");
                sb.AppendLine("        .info-section {");
                sb.AppendLine("            background-color: #f8f9fa;");
                sb.AppendLine("            padding: 20px;");
                sb.AppendLine("            border-radius: 8px;");
                sb.AppendLine("            margin-bottom: 30px;");
                sb.AppendLine("            border-left: 4px solid #3498db;");
                sb.AppendLine("        }");
                sb.AppendLine("        .info-item {");
                sb.AppendLine("            margin-bottom: 8px;");
                sb.AppendLine("            display: flex;");
                sb.AppendLine("        }");
                sb.AppendLine("        .info-label {");
                sb.AppendLine("            font-weight: bold;");
                sb.AppendLine("            min-width: 180px;");
                sb.AppendLine("            color: #2c3e50;");
                sb.AppendLine("        }");
                sb.AppendLine("        .stats-section {");
                sb.AppendLine("            display: grid;");
                sb.AppendLine("            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));");
                sb.AppendLine("            gap: 20px;");
                sb.AppendLine("            margin-bottom: 30px;");
                sb.AppendLine("        }");
                sb.AppendLine("        .stat-card {");
                sb.AppendLine("            background-color: white;");
                sb.AppendLine("            padding: 20px;");
                sb.AppendLine("            border-radius: 8px;");
                sb.AppendLine("            box-shadow: 0 2px 8px rgba(0,0,0,0.1);");
                sb.AppendLine("            text-align: center;");
                sb.AppendLine("            border-top: 4px solid;");
                sb.AppendLine("        }");
                sb.AppendLine("        .stat-card.total { border-top-color: #3498db; }");
                sb.AppendLine("        .stat-card.named { border-top-color: #2ecc71; }");
                sb.AppendLine("        .stat-card.critical { border-top-color: #e74c3c; }");
                sb.AppendLine("        .stat-card.real { border-top-color: #f39c12; }");
                sb.AppendLine("        .stat-card.all-duplicates { border-top-color: #9b59b6; }");
                sb.AppendLine("        .stat-number {");
                sb.AppendLine("            font-size: 36px;");
                sb.AppendLine("            font-weight: bold;");
                sb.AppendLine("            margin: 10px 0;");
                sb.AppendLine("        }");
                sb.AppendLine("        .stat-label {");
                sb.AppendLine("            font-size: 14px;");
                sb.AppendLine("            color: #7f8c8d;");
                sb.AppendLine("        }");
                sb.AppendLine("        .duplicate-group {");
                sb.AppendLine("            margin-bottom: 40px;");
                sb.AppendLine("            background-color: white;");
                sb.AppendLine("            border-radius: 8px;");
                sb.AppendLine("            overflow: hidden;");
                sb.AppendLine("            box-shadow: 0 2px 10px rgba(0,0,0,0.1);");
                sb.AppendLine("        }");
                sb.AppendLine("        .group-header {");
                sb.AppendLine("            background-color: #34495e;");
                sb.AppendLine("            color: white;");
                sb.AppendLine("            padding: 15px 20px;");
                sb.AppendLine("            font-size: 18px;");
                sb.AppendLine("        }");
                sb.AppendLine("        .group-info {");
                sb.AppendLine("            padding: 15px 20px;");
                sb.AppendLine("            background-color: #f8f9fa;");
                sb.AppendLine("            border-bottom: 1px solid #dee2e6;");
                sb.AppendLine("        }");
                sb.AppendLine("        .group-count {");
                sb.AppendLine("            display: inline-block;");
                sb.AppendLine("            background-color: #e74c3c;");
                sb.AppendLine("            color: white;");
                sb.AppendLine("            padding: 2px 10px;");
                sb.AppendLine("            border-radius: 20px;");
                sb.AppendLine("            font-size: 14px;");
                sb.AppendLine("            margin-left: 10px;");
                sb.AppendLine("        }");
                sb.AppendLine("        .objects-table {");
                sb.AppendLine("            width: 100%;");
                sb.AppendLine("            border-collapse: collapse;");
                sb.AppendLine("        }");
                sb.AppendLine("        .objects-table th {");
                sb.AppendLine("            background-color: #2c3e50;");
                sb.AppendLine("            color: white;");
                sb.AppendLine("            padding: 12px 15px;");
                sb.AppendLine("            text-align: left;");
                sb.AppendLine("            font-weight: 600;");
                sb.AppendLine("        }");
                sb.AppendLine("        .objects-table td {");
                sb.AppendLine("            padding: 10px 15px;");
                sb.AppendLine("            border-bottom: 1px solid #dee2e6;");
                sb.AppendLine("        }");
                sb.AppendLine("        .objects-table tr:nth-child(even) {");
                sb.AppendLine("            background-color: #f8f9fa;");
                sb.AppendLine("        }");
                sb.AppendLine("        .objects-table tr:hover {");
                sb.AppendLine("            background-color: #e9ecef;");
                sb.AppendLine("        }");
                sb.AppendLine("        .id-cell {");
                sb.AppendLine("            font-weight: bold;");
                sb.AppendLine("            color: #2980b9;");
                sb.AppendLine("            font-family: monospace;");
                sb.AppendLine("        }");
                sb.AppendLine("        .no-duplicates {");
                sb.AppendLine("            text-align: center;");
                sb.AppendLine("            padding: 40px;");
                sb.AppendLine("            color: #27ae60;");
                sb.AppendLine("            font-size: 18px;");
                sb.AppendLine("            background-color: #f8f9fa;");
                sb.AppendLine("            border-radius: 8px;");
                sb.AppendLine("            margin: 20px 0;");
                sb.AppendLine("        }");
                sb.AppendLine("        .footer {");
                sb.AppendLine("            text-align: center;");
                sb.AppendLine("            margin-top: 30px;");
                sb.AppendLine("            padding-top: 20px;");
                sb.AppendLine("            border-top: 1px solid #dee2e6;");
                sb.AppendLine("            color: #7f8c8d;");
                sb.AppendLine("            font-size: 12px;");
                sb.AppendLine("        }");
                sb.AppendLine("        .badge {");
                sb.AppendLine("            display: inline-block;");
                sb.AppendLine("            padding: 3px 8px;");
                sb.AppendLine("            border-radius: 4px;");
                sb.AppendLine("            font-size: 12px;");
                sb.AppendLine("            font-weight: bold;");
                sb.AppendLine("        }");
                sb.AppendLine("        .badge-critical {");
                sb.AppendLine("            background-color: #ffebee;");
                sb.AppendLine("            color: #c62828;");
                sb.AppendLine("        }");
                sb.AppendLine("        .badge-real {");
                sb.AppendLine("            background-color: #fff3e0;");
                sb.AppendLine("            color: #ef6c00;");
                sb.AppendLine("        }");
                sb.AppendLine("        @media print {");
                sb.AppendLine("            body { background-color: white; }");
                sb.AppendLine("            .container { box-shadow: none; }");
                sb.AppendLine("        }");
                sb.AppendLine("    </style>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");
                sb.AppendLine("    <div class='container'>");

                // ШАПКА ОТЧЕТА
                sb.AppendLine("        <div class='header'>");
                sb.AppendLine($"            <h1>📊 Отчет о поиске дубликатов в папке</h1>");
                sb.AppendLine($"            <div class='date'>Дата создания: {DateTime.Now:dd.MM.yyyy HH:mm:ss}</div>");
                sb.AppendLine("        </div>");

                // ИНФОРМАЦИЯ О ПАПКЕ
                sb.AppendLine("        <div class='info-section'>");
                sb.AppendLine("            <div class='info-item'><span class='info-label'>📁 Папка:</span> " + EscapeHtml(folderInfo.Version) + "</div>");
                sb.AppendLine("            <div class='info-item'><span class='info-label'>🆔 ID папки:</span> " + folderInfo.Id + "</div>");
                sb.AppendLine("            <div class='info-item'><span class='info-label'>📝 Тип:</span> " + EscapeHtml(folderInfo.Type) + "</div>");
                sb.AppendLine("            <div class='info-item'><span class='info-label'>📋 Состояние:</span> " + EscapeHtml(folderInfo.State) + "</div>");
                sb.AppendLine("        </div>");

                // СТАТИСТИКА
                sb.AppendLine("        <div class='stats-section'>");
                sb.AppendLine("            <div class='stat-card total'>");
                sb.AppendLine($"                <div class='stat-number'>{totalObjectsCount}</div>");
                sb.AppendLine("                <div class='stat-label'>Всего объектов в папке</div>");
                sb.AppendLine("            </div>");
                sb.AppendLine("            <div class='stat-card named'>");
                sb.AppendLine($"                <div class='stat-number'>{objectsWithNamesCount}</div>");
                sb.AppendLine("                <div class='stat-label'>Объектов с заполненными названиями</div>");
                sb.AppendLine("            </div>");
                sb.AppendLine("            <div class='stat-card critical'>");
                sb.AppendLine($"                <div class='stat-number'>{fullDuplicates.Count}</div>");
                sb.AppendLine("                <div class='stat-label'>Групп критических дубликатов</div>");
                sb.AppendLine("                <div style='font-size:12px; margin-top:5px;'>(Название + Тип + Версия)</div>");
                sb.AppendLine("            </div>");
                sb.AppendLine("            <div class='stat-card real'>");
                sb.AppendLine($"                <div class='stat-number'>{nameTypeDuplicates.Count}</div>");
                sb.AppendLine("                <div class='stat-label'>Групп реальных дубликатов</div>");
                sb.AppendLine("                <div style='font-size:12px; margin-top:5px;'>(Название + Тип)</div>");
                sb.AppendLine("            </div>");
                sb.AppendLine("            <div class='stat-card all-duplicates'>");
                sb.AppendLine($"                <div class='stat-number'>{totalRealDuplicates}</div>");
                sb.AppendLine("                <div class='stat-label'>Всего групп дубликатов</div>");
                sb.AppendLine("            </div>");
                sb.AppendLine("        </div>");

                // ДУБЛИКАТЫ
                if (totalRealDuplicates > 0)
                {
                    // Критические дубликаты
                    if (fullDuplicates.Count > 0)
                    {
                        sb.AppendLine("        <div class='duplicate-group'>");
                        sb.AppendLine("            <div class='group-header'>");
                        sb.AppendLine("                ⚠️ Критические дубликаты");
                        sb.AppendLine("                <span class='group-count'>" + fullDuplicates.Count + " групп</span>");
                        sb.AppendLine("            </div>");
                        sb.AppendLine("            <div class='group-info'>");
                        sb.AppendLine("                Полное совпадение: Название + Тип + Версия<br>");
                        sb.AppendLine("                <span class='badge badge-critical'>ВНИМАНИЕ: Эти объекты полностью идентичны!</span>");
                        sb.AppendLine("            </div>");

                        int groupNum = 1;
                        foreach (var group in fullDuplicates.OrderByDescending(g => g.Count()))
                        {
                            string[] parts = group.Key.Split('|');

                            sb.AppendLine("            <div style='padding:20px; border-bottom:2px solid #e74c3c;'>");
                            sb.AppendLine($"                <h3 style='color:#c62828; margin-bottom:10px;'>Группа #{groupNum++}: \"{EscapeHtml(parts[0])}\"</h3>");
                            sb.AppendLine($"                <div style='margin-bottom:15px; color:#555;'>");
                            sb.AppendLine($"                    <strong>Тип:</strong> {EscapeHtml(parts[1])}<br>");
                            sb.AppendLine($"                    <strong>Версия:</strong> {EscapeHtml(parts[2])}<br>");
                            sb.AppendLine($"                    <strong>Количество:</strong> <span style='color:#e74c3c; font-weight:bold;'>{group.Count()} объектов</span>");
                            sb.AppendLine("                </div>");

                            sb.AppendLine("                <table class='objects-table'>");
                            sb.AppendLine("                    <thead>");
                            sb.AppendLine("                        <tr>");
                            sb.AppendLine("                            <th>ID</th>");
                            sb.AppendLine("                            <th>Название</th>");
                            sb.AppendLine("                            <th>Тип</th>");
                            sb.AppendLine("                            <th>Версия</th>");
                            sb.AppendLine("                            <th>Состояние</th>");
                            sb.AppendLine("                        </tr>");
                            sb.AppendLine("                    </thead>");
                            sb.AppendLine("                    <tbody>");

                            foreach (var obj in group.OrderBy(o => o.Id))
                            {
                                sb.AppendLine("                        <tr>");
                                sb.AppendLine($"                            <td class='id-cell'>{obj.Id}</td>");
                                sb.AppendLine($"                            <td>{EscapeHtml(obj.Name)}</td>");
                                sb.AppendLine($"                            <td>{EscapeHtml(obj.Type)}</td>");
                                sb.AppendLine($"                            <td>{EscapeHtml(obj.Version)}</td>");
                                sb.AppendLine($"                            <td>{EscapeHtml(obj.State)}</td>");
                                sb.AppendLine("                        </tr>");
                            }

                            sb.AppendLine("                    </tbody>");
                            sb.AppendLine("                </table>");
                            sb.AppendLine("            </div>");
                        }
                        sb.AppendLine("        </div>");
                    }

                    // Реальные дубликаты
                    if (nameTypeDuplicates.Count > 0)
                    {
                        sb.AppendLine("        <div class='duplicate-group'>");
                        sb.AppendLine("            <div class='group-header'>");
                        sb.AppendLine("                🔍 Реальные дубликаты");
                        sb.AppendLine("                <span class='group-count'>" + nameTypeDuplicates.Count + " групп</span>");
                        sb.AppendLine("            </div>");
                        sb.AppendLine("            <div class='group-info'>");
                        sb.AppendLine("                Совпадение: Название + Тип<br>");
                        sb.AppendLine("                <span class='badge badge-real'>Проверьте версии этих объектов</span>");
                        sb.AppendLine("            </div>");

                        int groupNum = 1;
                        foreach (var group in nameTypeDuplicates.OrderByDescending(g => g.Count()))
                        {
                            string[] parts = group.Key.Split('|');

                            sb.AppendLine("            <div style='padding:20px; border-bottom:2px solid #f39c12;'>");
                            sb.AppendLine($"                <h3 style='color:#e67e22; margin-bottom:10px;'>Группа #{groupNum++}: \"{EscapeHtml(parts[0])}\"</h3>");
                            sb.AppendLine($"                <div style='margin-bottom:15px; color:#555;'>");
                            sb.AppendLine($"                    <strong>Тип:</strong> {EscapeHtml(parts[1])}<br>");
                            sb.AppendLine($"                    <strong>Количество:</strong> <span style='color:#f39c12; font-weight:bold;'>{group.Count()} объектов</span>");
                            sb.AppendLine("                </div>");

                            sb.AppendLine("                <table class='objects-table'>");
                            sb.AppendLine("                    <thead>");
                            sb.AppendLine("                        <tr>");
                            sb.AppendLine("                            <th>ID</th>");
                            sb.AppendLine("                            <th>Название</th>");
                            sb.AppendLine("                            <th>Тип</th>");
                            sb.AppendLine("                            <th>Версия</th>");
                            sb.AppendLine("                            <th>Состояние</th>");
                            sb.AppendLine("                        </tr>");
                            sb.AppendLine("                    </thead>");
                            sb.AppendLine("                    <tbody>");

                            foreach (var obj in group.OrderBy(o => o.Version).ThenBy(o => o.Id))
                            {
                                sb.AppendLine("                        <tr>");
                                sb.AppendLine($"                            <td class='id-cell'>{obj.Id}</td>");
                                sb.AppendLine($"                            <td>{EscapeHtml(obj.Name)}</td>");
                                sb.AppendLine($"                            <td>{EscapeHtml(obj.Type)}</td>");
                                sb.AppendLine($"                            <td>{EscapeHtml(obj.Version)}</td>");
                                sb.AppendLine($"                            <td>{EscapeHtml(obj.State)}</td>");
                                sb.AppendLine("                        </tr>");
                            }

                            sb.AppendLine("                    </tbody>");
                            sb.AppendLine("                </table>");
                            sb.AppendLine("            </div>");
                        }
                        sb.AppendLine("        </div>");
                    }
                }
                else
                {
                    sb.AppendLine("        <div class='no-duplicates'>");
                    sb.AppendLine("            ✅ Реальные дубликаты не найдены");
                    sb.AppendLine("        </div>");
                }

                // ПОДВАЛ
                sb.AppendLine("        <div class='footer'>");
                sb.AppendLine("            Отчет сгенерирован плагином DeepDuplicateFinder<br>");
                sb.AppendLine($"            {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine("        </div>");

                sb.AppendLine("    </div>");
                sb.AppendLine("</body>");
                sb.AppendLine("</html>");

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

                // Автоматическое открытие файла в браузере
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    // Если не удалось открыть, просто игнорируем
                    System.Diagnostics.Debug.WriteLine($"Не удалось открыть HTML файл: {ex.Message}");
                }

                return filePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка сохранения HTML отчета: {ex.Message}");
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

    // Класс для хранения информации об объекте
    public class ObjectInfo
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string Version { get; set; } = "";
        public string State { get; set; } = "";
    }
}