using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace DeepDuplicateFinder
{
    public class MaterialDuplicatesHtmlReportGenerator
    {
        public string CreateHtmlReport(ObjectInfo folderInfo,
                             int totalObjectsCount,
                             int totalDetailsCount,
                             int totalMaterialsFound,
                             int totalDuplicateMaterials,
                             List<DetailWithMaterialDuplicates> detailsWithDuplicates,
                             string materialTypeName = "Материал по КД")
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string safeName = ReplaceInvalidChars(folderInfo.Name);
                string fileName = $"material_duplicates_report_{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.html";
                string filePath = Path.Combine(desktop, fileName);
                var sb = new StringBuilder();
                // Начало HTML документа
                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("<html lang='ru'>");
                sb.AppendLine("<head>");
                sb.AppendLine(" <meta charset='UTF-8'>");
                sb.AppendLine(" <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
                sb.AppendLine($" <title>Отчет о дубликатах материалов - {EscapeHtml(folderInfo.Name)}</title>");
                sb.AppendLine(" <style>");
                sb.AppendLine(" * {");
                sb.AppendLine(" margin: 0;");
                sb.AppendLine(" padding: 0;");
                sb.AppendLine(" box-sizing: border-box;");
                sb.AppendLine(" font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;");
                sb.AppendLine(" }");
                sb.AppendLine(" body {");
                sb.AppendLine(" background-color: #f5f5f5;");
                sb.AppendLine(" color: #333;");
                sb.AppendLine(" line-height: 1.6;");
                sb.AppendLine(" padding: 20px;");
                sb.AppendLine(" }");
                sb.AppendLine(" .container {");
                sb.AppendLine(" max-width: 1200px;");
                sb.AppendLine(" margin: 0 auto;");
                sb.AppendLine(" background-color: white;");
                sb.AppendLine(" border-radius: 10px;");
                sb.AppendLine(" box-shadow: 0 2px 15px rgba(0,0,0,0.1);");
                sb.AppendLine(" padding: 30px;");
                sb.AppendLine(" }");
                sb.AppendLine(" .header {");
                sb.AppendLine(" text-align: center;");
                sb.AppendLine(" margin-bottom: 30px;");
                sb.AppendLine(" padding-bottom: 20px;");
                sb.AppendLine(" border-bottom: 3px solid #4CAF50;");
                sb.AppendLine(" }");
                sb.AppendLine(" .date {");
                sb.AppendLine(" color: #666;");
                sb.AppendLine(" font-size: 14px;");
                sb.AppendLine(" margin-top: 5px;");
                sb.AppendLine(" }");
                sb.AppendLine(" .info-section {");
                sb.AppendLine(" background-color: #f8f9fa;");
                sb.AppendLine(" padding: 20px;");
                sb.AppendLine(" border-radius: 8px;");
                sb.AppendLine(" margin-bottom: 30px;");
                sb.AppendLine(" border-left: 4px solid #3498db;");
                sb.AppendLine(" }");
                sb.AppendLine(" .info-item {");
                sb.AppendLine(" margin-bottom: 10px;");
                sb.AppendLine(" font-size: 16px;");
                sb.AppendLine(" }");
                sb.AppendLine(" .info-label {");
                sb.AppendLine(" font-weight: bold;");
                sb.AppendLine(" margin-right: 5px;");
                sb.AppendLine(" }");
                sb.AppendLine(" .stats-section {");
                sb.AppendLine(" display: grid;");
                sb.AppendLine(" grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));");
                sb.AppendLine(" gap: 20px;");
                sb.AppendLine(" margin-bottom: 30px;");
                sb.AppendLine(" }");
                sb.AppendLine(" .stat-card {");
                sb.AppendLine(" background-color: white;");
                sb.AppendLine(" padding: 20px;");
                sb.AppendLine(" border-radius: 8px;");
                sb.AppendLine(" box-shadow: 0 2px 8px rgba(0,0,0,0.1);");
                sb.AppendLine(" text-align: center;");
                sb.AppendLine(" border-top: 4px solid;");
                sb.AppendLine(" }");
                sb.AppendLine(" .stat-card.total { border-top-color: #3498db; }");
                sb.AppendLine(" .stat-card.details { border-top-color: #2ecc71; }");
                sb.AppendLine(" .stat-card.materials { border-top-color: #f39c12; }");
                sb.AppendLine(" .stat-card.duplicates { border-top-color: #e74c3c; }");
                sb.AppendLine(" .stat-card.with-duplicates { border-top-color: #9b59b6; }");
                sb.AppendLine(" .stat-number {");
                sb.AppendLine(" font-size: 36px;");
                sb.AppendLine(" font-weight: bold;");
                sb.AppendLine(" margin: 10px 0;");
                sb.AppendLine(" }");
                sb.AppendLine(" .stat-label {");
                sb.AppendLine(" font-size: 16px;");
                sb.AppendLine(" color: #666;");
                sb.AppendLine(" }");
                sb.AppendLine(" .detail-section {");
                sb.AppendLine(" margin-bottom: 40px;");
                sb.AppendLine(" background-color: white;");
                sb.AppendLine(" border-radius: 8px;");
                sb.AppendLine(" overflow: hidden;");
                sb.AppendLine(" box-shadow: 0 2px 10px rgba(0,0,0,0.1);");
                sb.AppendLine(" }");
                sb.AppendLine(" .detail-header {");
                sb.AppendLine(" background-color: #34495e;");
                sb.AppendLine(" color: white;");
                sb.AppendLine(" padding: 15px 20px;");
                sb.AppendLine(" font-size: 18px;");
                sb.AppendLine(" }");
                sb.AppendLine(" .detail-info {");
                sb.AppendLine(" padding: 15px 20px;");
                sb.AppendLine(" background-color: #f8f9fa;");
                sb.AppendLine(" border-bottom: 1px solid #dee2e6;");
                sb.AppendLine(" }");
                sb.AppendLine(" .material-group {");
                sb.AppendLine(" margin: 20px;");
                sb.AppendLine(" padding: 15px;");
                sb.AppendLine(" background-color: #fff3e0;");
                sb.AppendLine(" border-radius: 6px;");
                sb.AppendLine(" border-left: 4px solid #ff9800;");
                sb.AppendLine(" }");
                sb.AppendLine(" .objects-table {");
                sb.AppendLine(" width: 100%;");
                sb.AppendLine(" border-collapse: collapse;");
                sb.AppendLine(" margin-top: 15px;");
                sb.AppendLine(" }");
                sb.AppendLine(" .objects-table th {");
                sb.AppendLine(" background-color: #2c3e50;");
                sb.AppendLine(" color: white;");
                sb.AppendLine(" padding: 12px 15px;");
                sb.AppendLine(" text-align: left;");
                sb.AppendLine(" font-weight: 600;");
                sb.AppendLine(" }");
                sb.AppendLine(" .objects-table td {");
                sb.AppendLine(" padding: 10px 15px;");
                sb.AppendLine(" border-bottom: 1px solid #dee2e6;");
                sb.AppendLine(" }");
                sb.AppendLine(" .no-duplicates {");
                sb.AppendLine(" text-align: center;");
                sb.AppendLine(" padding: 40px;");
                sb.AppendLine(" color: #27ae60;");
                sb.AppendLine(" font-size: 18px;");
                sb.AppendLine(" background-color: #f8f9fa;");
                sb.AppendLine(" border-radius: 8px;");
                sb.AppendLine(" margin: 20px 0;");
                sb.AppendLine(" }");
                sb.AppendLine(" </style>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");
                sb.AppendLine(" <div class='container'>");
                // ШАПКА ОТЧЕТА
                sb.AppendLine(" <div class='header'>");
                sb.AppendLine(" <h1>🔍 Отчет о дубликатах материалов</h1>");
                sb.AppendLine($" <div class='date'>Дата создания: {DateTime.Now:dd.MM.yyyy HH:mm:ss}</div>");
                sb.AppendLine(" </div>");
                // ИНФОРМАЦИЯ О ПАПКЕ
                sb.AppendLine(" <div class='info-section'>");
                sb.AppendLine(" <div class='info-item'><span class='info-label'>📁 Папка:</span> " + EscapeHtml(folderInfo.Name) + "</div>");
                sb.AppendLine(" <div class='info-item'><span class='info-label'>🆔 ID папки:</span> " + folderInfo.Id + "</div>");
                sb.AppendLine(" <div class='info-item'><span class='info-label'>🔍 Тип материалов:</span> " + EscapeHtml(materialTypeName) + "</div>");
                sb.AppendLine(" <div class='info-item'><span class='info-label'>📝 Пояснение:</span> Ищем дубликаты материалов ВНУТРИ объектов (не путать с одинаковыми материалами в разных объектах)</div>");
                sb.AppendLine(" </div>");
                // СТАТИСТИКА
                sb.AppendLine(" <div class='stats-section'>");
                sb.AppendLine(" <div class='stat-card total'>");
                sb.AppendLine($" <div class='stat-number'>{totalObjectsCount}</div>");
                sb.AppendLine(" <div class='stat-label'>Всего объектов в папке</div>");
                sb.AppendLine(" </div>");
                sb.AppendLine(" <div class='stat-card details'>");
                sb.AppendLine($" <div class='stat-number'>{totalDetailsCount}</div>");
                sb.AppendLine(" <div class='stat-label'>Всего деталей</div>");
                sb.AppendLine(" </div>");
                sb.AppendLine(" <div class='stat-card materials'>");
                sb.AppendLine($" <div class='stat-number'>{totalMaterialsFound}</div>");
                sb.AppendLine(" <div class='stat-label'>Всего связей с материалами</div>");
                sb.AppendLine(" </div>");
                sb.AppendLine(" <div class='stat-card duplicates'>");
                sb.AppendLine($" <div class='stat-number'>{totalDuplicateMaterials}</div>");
                sb.AppendLine(" <div class='stat-label'>Всего дубликатов материалов</div>");
                sb.AppendLine(" </div>");
                sb.AppendLine(" <div class='stat-card with-duplicates'>");
                sb.AppendLine($" <div class='stat-number'>{detailsWithDuplicates.Count}</div>");
                sb.AppendLine(" <div class='stat-label'>Деталей с дубликатами</div>");
                sb.AppendLine(" </div>");
                sb.AppendLine(" </div>");
                // ОБЪЕКТЫ С ДУБЛИКАТАМИ МАТЕРИАЛОВ
                if (detailsWithDuplicates.Count > 0)
                {
                    sb.AppendLine($" <h2>Найдено {detailsWithDuplicates.Count} объектов с дубликатами материалов:</h2>");
                    int objectCounter = 1;
                    foreach (var objWithDuplicates in detailsWithDuplicates)
                    {
                        var parentObject = objWithDuplicates.Detail;
                        sb.AppendLine(" <div class='detail-section'>");
                        sb.AppendLine(" <div class='detail-header'>");
                        sb.AppendLine($" 📦 Объект #{objectCounter++}: {EscapeHtml(parentObject.Name)}");
                        sb.AppendLine(" </div>");
                        sb.AppendLine(" <div class='detail-info'>");
                        sb.AppendLine($" <strong>ID:</strong> {parentObject.Id}<br>");
                        sb.AppendLine($" <strong>Тип:</strong> {EscapeHtml(parentObject.Type)}<br>");
                        sb.AppendLine($" <strong>Версия:</strong> {EscapeHtml(parentObject.Version)}<br>");
                        sb.AppendLine($" <strong>Состояние:</strong> {EscapeHtml(parentObject.State)}");
                        sb.AppendLine(" </div>");
                        int groupCounter = 1;
                        foreach (var materialGroup in objWithDuplicates.MaterialGroups)
                        {
                            var firstMaterial = materialGroup.Materials.FirstOrDefault();
                            if (firstMaterial == null)
                                continue;
                            sb.AppendLine(" <div class='material-group'>");
                            sb.AppendLine($" <h3>Дубликат материала (ID: {materialGroup.Key})</h3>");
                            sb.AppendLine($" <p><strong>Название:</strong> {EscapeHtml(firstMaterial.Name)}</p>");
                            sb.AppendLine($" <p><strong>Тип:</strong> {EscapeHtml(firstMaterial.Type)}</p>");
                            sb.AppendLine($" <p><strong>Количество связей в детали:</strong> {materialGroup.Materials.Count}</p>");
                            sb.AppendLine(" <table class='objects-table'>");
                            sb.AppendLine(" <thead>");
                            sb.AppendLine(" <tr>");
                            sb.AppendLine(" <th>ID материала</th>");
                            sb.AppendLine(" <th>Название</th>");
                            sb.AppendLine(" <th>Тип</th>");
                            sb.AppendLine(" <th>Версия</th>");
                            sb.AppendLine(" <th>Состояние</th>");
                            sb.AppendLine(" <th>Количество связей</th>");
                            sb.AppendLine(" </tr>");
                            sb.AppendLine(" </thead>");
                            sb.AppendLine(" <tbody>");
                            // Выводим только одну строку для каждого дублируемого материала
                            sb.AppendLine(" <tr>");
                            sb.AppendLine($" <td>{materialGroup.Key}</td>");
                            sb.AppendLine($" <td>{EscapeHtml(firstMaterial.Name)}</td>");
                            sb.AppendLine($" <td>{EscapeHtml(firstMaterial.Type)}</td>");
                            sb.AppendLine($" <td>{EscapeHtml(firstMaterial.Version)}</td>");
                            sb.AppendLine($" <td>{EscapeHtml(firstMaterial.State)}</td>");
                            sb.AppendLine($" <td>{materialGroup.Materials.Count}</td>");
                            sb.AppendLine(" </tr>");
                            sb.AppendLine(" </tbody>");
                            sb.AppendLine(" </table>");
                            sb.AppendLine(" </div>");
                        }
                        sb.AppendLine(" </div>");
                    }
                }
                else
                {
                    sb.AppendLine(" <div class='no-duplicates'>");
                    sb.AppendLine(" ✅ Дубликаты материалов в объектах не найдены");
                    sb.AppendLine(" <p style='margin-top:10px; color:#666;'>Одинаковые материалы в разных объектах не считаются дубликатами.</p>");
                    sb.AppendLine(" </div>");
                }
                sb.AppendLine(" </div>");
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
            return WebUtility.HtmlEncode(text);
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