using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Ascon.Plm.Loodsman.PluginSDK;

namespace DeepDuplicateFinder
{
    [LoodsmanPlugin]
    public class DeepDuplicateFinder : ILoodsmanNetPlugin
    {
        private Dictionary<int, string> _typeDictionary = new Dictionary<int, string>();

        public void PluginLoad()
        {
            // Логирование удалено
        }

        public void PluginUnload()
        {
            // Логирование удалено
        }

        public void OnConnectToDb(INetPluginCall call)
        {
            // Логирование удалено
        }

        public void OnCloseDb()
        {
            // Логирование удалено
        }

        public void BindMenu(IMenuDefinition menu)
        {
            menu.AddMenuItem("Найти дубликаты в папке", FindDuplicatesInFolder, call => true);
        }

        private void FindDuplicatesInFolder(INetPluginCall call)
        {
            try
            {
                // 1. Получаем выделенные ID
                int selectedId = GetSelectedId(call);

                if (selectedId == 0)
                {
                    return;
                }

                // 2. Устанавливаем XML формат для получения данных
                try
                {
                    call.RunMethod("SetFormat", new object[] { "xml" });
                }
                catch (Exception ex)
                {
                    return;
                }

                // 3. Получаем информацию о папке
                var folderInfo = GetObjectInfo(call, selectedId);

                // 4. Получаем словарь типов для замены ID на названия
                _typeDictionary = GetTypeDictionary(call);

                // 5. Получаем объекты в папке
                List<ObjectInfo> allObjects = GetObjectsInFolder(call, selectedId);

                if (allObjects.Count == 0)
                {
                    // Возвращаем бинарный формат
                    try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
                    return;
                }

                // Заменяем ID типов на названия
                foreach (var obj in allObjects)
                {
                    if (int.TryParse(obj.Type, out int typeId) && _typeDictionary.ContainsKey(typeId))
                    {
                        obj.Type = _typeDictionary[typeId];
                    }
                }

                // 6. Ищем дубликаты
                // Фильтруем объекты с названиями
                var objectsWithNames = allObjects
                    .Where(o => !string.IsNullOrEmpty(o.Name) && o.Name != $"Объект_{o.Id}")
                    .ToList();

                // 6.1. Критические дубликаты по полному совпадению: Название + Тип + Версия
                var fullDuplicates = objectsWithNames
                    .Where(o => !string.IsNullOrEmpty(o.Type) && !string.IsNullOrEmpty(o.Version))
                    .GroupBy(o => $"{o.Name.Trim().ToLower()}|{o.Type.Trim().ToLower()}|{o.Version.Trim().ToLower()}")
                    .Where(g => g.Count() > 1)
                    .OrderByDescending(g => g.Count())
                    .ToList();

                // 6.2. Реальные дубликаты по названию + типу (одинаковое название и тип - это ДУБЛИКАТ)
                var nameTypeDuplicates = objectsWithNames
                    .Where(o => !string.IsNullOrEmpty(o.Type))
                    .GroupBy(o => $"{o.Name.Trim().ToLower()}|{o.Type.Trim().ToLower()}")
                    .Where(g => g.Count() > 1)
                    .OrderByDescending(g => g.Count())
                    .ToList();

                int totalRealDuplicates = fullDuplicates.Count + nameTypeDuplicates.Count;

                // Возвращаем бинарный формат
                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }

                // 7. Сохраняем отчет
                string reportPath = SaveCombinedReport(folderInfo, allObjects.Count, objectsWithNames.Count,
                    fullDuplicates, nameTypeDuplicates, totalRealDuplicates);
            }
            catch (Exception ex)
            {
                // Пытаемся вернуть бинарный формат в случае ошибки
                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
            }
        }

        private string SaveCombinedReport(ObjectInfo folderInfo,
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
                string fileName = $"duplicates_report_{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string filePath = Path.Combine(desktop, fileName);

                var sb = new StringBuilder();

                // ШАПКА ОТЧЕТА
                sb.AppendLine(new string('=', 70));
                sb.AppendLine("ОТЧЕТ О ПОИСКЕ ДУБЛИКАТОВ В ПАПКЕ");
                sb.AppendLine(new string('=', 70));
                sb.AppendLine($"Дата: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
                sb.AppendLine($"Название папка: {folderInfo.Version} (ID: {folderInfo.Id})");
                sb.AppendLine($"Тип: {folderInfo.Type}");
                sb.AppendLine($"Состояние: {folderInfo.State}");
                sb.AppendLine();
                sb.AppendLine(new string('-', 70));
                sb.AppendLine();

                sb.AppendLine($"ВСЕГО ОБЪЕКТОВ В ПАПКЕ: {totalObjectsCount}");
                sb.AppendLine($"Объектов с заполненными названиями: {objectsWithNamesCount}");
                sb.AppendLine();

                // Информация о дубликатах
                sb.AppendLine("СТАТИСТИКА ДУБЛИКАТОВ:");
                sb.AppendLine($"• Критические (Название+Тип+Версия): {fullDuplicates.Count}");
                sb.AppendLine($"• Реальные дубликаты (Название+Тип): {nameTypeDuplicates.Count}");
                sb.AppendLine($"• ВСЕГО РЕАЛЬНЫХ ДУБЛИКАТОВ: {totalRealDuplicates}");
                sb.AppendLine();

                // Если есть дубликаты, показываем их
                if (totalRealDuplicates > 0)
                {
                    if (fullDuplicates.Count > 0)
                    {
                        sb.AppendLine("КРИТИЧЕСКИЕ ДУБЛИКАТЫ (полное совпадение):");
                        sb.AppendLine(new string('-', 50));
                        int groupNum = 1;
                        foreach (var group in fullDuplicates.OrderByDescending(g => g.Count()))
                        {
                            string[] parts = group.Key.Split('|');
                            sb.AppendLine($"ГРУППА #{groupNum++}: '{parts[0]}' (Тип: '{parts[1]}', Версия: '{parts[2]}')");
                            sb.AppendLine($"Количество: {group.Count()} объектов");
                            sb.AppendLine();

                            foreach (var obj in group.OrderBy(o => o.Id))
                            {
                                sb.AppendLine($"  ID: {obj.Id}");
                                sb.AppendLine($"    Название: {obj.Name}");
                                sb.AppendLine($"    Тип: {obj.Type}");
                                sb.AppendLine($"    Версия: {obj.Version}");
                                sb.AppendLine($"    Состояние: {obj.State}");
                                sb.AppendLine();
                            }
                            sb.AppendLine(new string('-', 50));
                            sb.AppendLine();
                        }
                    }

                    if (nameTypeDuplicates.Count > 0)
                    {
                        sb.AppendLine("РЕАЛЬНЫЕ ДУБЛИКАТЫ (одинаковое название, тип и версия):");
                        sb.AppendLine(new string('-', 50));
                        int groupNum = 1;
                        foreach (var group in nameTypeDuplicates.OrderByDescending(g => g.Count()))
                        {
                            string[] parts = group.Key.Split('|');
                            sb.AppendLine($"ГРУППА #{groupNum++}: '{parts[0]}' (Тип: '{parts[1]}')");
                            sb.AppendLine($"Количество: {group.Count()} объектов");
                            sb.AppendLine();

                            foreach (var obj in group.OrderBy(o => o.Version).ThenBy(o => o.Id))
                            {
                                sb.AppendLine($"  ID: {obj.Id}");
                                sb.AppendLine($"    Название: {obj.Name}");
                                sb.AppendLine($"    Тип: {obj.Type}");
                                sb.AppendLine($"    Версия: {obj.Version}");
                                sb.AppendLine($"    Состояние: {obj.State}");
                                sb.AppendLine();
                            }
                            sb.AppendLine(new string('-', 50));
                            sb.AppendLine();
                        }
                    }
                }
                else
                {
                    sb.AppendLine("РЕАЛЬНЫЕ ДУБЛИКАТЫ НЕ НАЙДЕНЫ");
                }

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                return filePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка сохранения отчета: {ex.Message}");
            }
        }

        private Dictionary<int, string> GetTypeDictionary(INetPluginCall call)
        {
            var typeDict = new Dictionary<int, string>();

            try
            {
                object result = call.RunMethod("GetTypeList", new object[0]);

                if (result is string xmlData && !string.IsNullOrEmpty(xmlData))
                {
                    var xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(xmlData);

                    XmlNodeList rows = xmlDoc.SelectNodes("//row");
                    if (rows == null || rows.Count == 0)
                        rows = xmlDoc.SelectNodes("//ROOT/rowset/row");

                    if (rows != null)
                    {
                        foreach (XmlNode row in rows)
                        {
                            int typeId = 0;
                            string typeName = "";

                            foreach (XmlAttribute attr in row.Attributes)
                            {
                                string attrName = attr.Name.ToUpper();
                                string attrValue = attr.Value;

                                if (attrName == "C0" || attrName == "_ID" || attrName == "ID")
                                {
                                    if (int.TryParse(attrValue, out int id))
                                        typeId = id;
                                }
                                else if (attrName == "C1" || attrName == "_TYPENAME" || attrName == "TYPENAME")
                                {
                                    typeName = attrValue;
                                }
                            }

                            if (typeId > 0 && !string.IsNullOrEmpty(typeName) && !typeDict.ContainsKey(typeId))
                            {
                                typeDict[typeId] = typeName;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка получения словаря типов: {ex.Message}");
            }

            return typeDict;
        }

        private int GetSelectedId(INetPluginCall call)
        {
            try
            {
                object result = call.RunMethod("CGetTreeSelectedIDs", new object[0]);
                if (result is string idsStr && !string.IsNullOrEmpty(idsStr))
                {
                    string[] ids = idsStr.Split(',');
                    if (ids.Length > 0 && int.TryParse(ids[0].Trim(), out int id) && id > 0)
                    {
                        return id;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка получения выделенных ID: {ex.Message}");
            }

            return 0;
        }

        private ObjectInfo GetObjectInfo(INetPluginCall call, int objectId)
        {
            try
            {
                object result = call.RunMethod("GetPropObjects", new object[]
                {
                    objectId.ToString(),  // stObjectList
                    0                     // inParams
                });

                if (result is string xmlData && !string.IsNullOrEmpty(xmlData))
                {
                    return ParseObjectInfoXml(xmlData, objectId);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка GetPropObjects: {ex.Message}");
            }

            return new ObjectInfo { Id = objectId, Name = $"Объект_{objectId}" };
        }

        private ObjectInfo ParseObjectInfoXml(string xmlData, int objectId)
        {
            try
            {
                var xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xmlData);

                var row = xmlDoc.SelectSingleNode("//row");
                if (row == null)
                    row = xmlDoc.SelectSingleNode("//ROOT/rowset/row");

                if (row != null)
                {
                    var info = new ObjectInfo { Id = objectId };

                    foreach (XmlAttribute attr in row.Attributes)
                    {
                        string attrName = attr.Name.ToUpper();
                        string attrValue = attr.Value;

                        if (attrName == "C1" || attrName == "_TYPE")
                        {
                            info.Type = attrValue;
                        }
                        else if (attrName == "C2" || attrName == "_PRODUCT")
                        {
                            info.Name = attrValue;
                        }
                        else if (attrName == "C3" || attrName == "_VERSION")
                        {
                            info.Version = attrValue;
                        }
                        else if (attrName == "C4" || attrName == "_STATE")
                        {
                            info.State = attrValue;
                        }
                    }

                    if (string.IsNullOrEmpty(info.Name))
                    {
                        info.Name = !string.IsNullOrEmpty(info.Type)
                            ? $"{info.Type} {info.Version}".Trim()
                            : $"Объект_{objectId}";
                    }

                    return info;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка парсинга объекта: {ex.Message}");
            }

            return new ObjectInfo { Id = objectId, Name = $"Объект_{objectId}" };
        }

        private List<ObjectInfo> GetObjectsInFolder(INetPluginCall call, int folderId)
        {
            try
            {
                // Метод поиска
                object result = call.RunMethod("GetAllLinkedObjects", new object[]
                {
                    folderId.ToString(),  // stIds
                    0                     // inParams
                });

                if (result is string xmlData && !string.IsNullOrEmpty(xmlData))
                {
                    return ParseTreeXml(xmlData);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка GetAllLinkedObjects: {ex.Message}");
            }

            return new List<ObjectInfo>();
        }

        private List<ObjectInfo> ParseTreeXml(string xmlData)
        {
            var objects = new List<ObjectInfo>();

            try
            {
                if (string.IsNullOrEmpty(xmlData))
                    return objects;

                var xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xmlData);

                // Пробуем разные пути к данным
                XmlNodeList rows = xmlDoc.SelectNodes("//row");
                if (rows == null || rows.Count == 0)
                    rows = xmlDoc.SelectNodes("//ROOT/rowset/row");

                if (rows != null && rows.Count > 0)
                {
                    foreach (XmlNode row in rows)
                    {
                        var info = new ObjectInfo();

                        foreach (XmlAttribute attr in row.Attributes)
                        {
                            string attrName = attr.Name.ToUpper();
                            string attrValue = attr.Value;

                            // Парсим ID
                            if (attrName == "C0" || attrName == "_ID_VERSION" || attrName == "ID" ||
                                attrName == "_ID" || attrName == "ID_VERSION" || attrName == "_ID_VERSION")
                            {
                                if (int.TryParse(attrValue, out int id))
                                    info.Id = id;
                            }
                            // Парсим Тип
                            else if (attrName == "C1" || attrName == "_TYPE" || attrName == "TYPE")
                            {
                                info.Type = attrValue;
                            }
                            // Парсим Название/Продукт
                            else if (attrName == "C2" || attrName == "C3" || attrName == "_PRODUCT" ||
                                     attrName == "NAME" || attrName == "_NAME" || attrName == "PRODUCT")
                            {
                                info.Name = attrValue;
                            }
                            // Парсим Версию
                            else if (attrName == "C4" || attrName == "_VERSION" || attrName == "VERSION")
                            {
                                info.Version = attrValue;
                            }
                            // Парсим Состояние
                            else if (attrName == "C5" || attrName == "_STATE" || attrName == "STATE")
                            {
                                info.State = attrValue;
                            }
                        }

                        if (info.Id > 0)
                        {
                            // Если имя пустое, создаем из типа и версии
                            if (string.IsNullOrEmpty(info.Name))
                            {
                                info.Name = !string.IsNullOrEmpty(info.Type)
                                    ? $"{info.Type} {info.Version}".Trim()
                                    : $"Объект_{info.Id}";
                            }
                            objects.Add(info);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка парсинга XML: {ex.Message}");
            }

            return objects;
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

        private class ObjectInfo
        {
            public int Id { get; set; }
            public string Name { get; set; } = "";
            public string Type { get; set; } = "";
            public string Version { get; set; } = "";
            public string State { get; set; } = "";
        }
    }
}