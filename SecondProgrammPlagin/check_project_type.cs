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
        private string _logFilePath;
        private Dictionary<int, string> _typeDictionary = new Dictionary<int, string>();

        public void PluginLoad()
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"duplicate_finder_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                _logFilePath = Path.Combine(desktop, fileName);

                File.WriteAllText(_logFilePath, $"=== DeepDuplicateFinder ЗАГРУЖЕН {DateTime.Now} ===\n\n", Encoding.UTF8);
            }
            catch { }
        }

        public void PluginUnload()
        {
            try
            {
                File.AppendAllText(_logFilePath, $"\n=== DeepDuplicateFinder ВЫГРУЖЕН {DateTime.Now} ===\n", Encoding.UTF8);
            }
            catch { }
        }

        public void OnConnectToDb(INetPluginCall call)
        {
            LogMessage("Подключение к базе данных");
        }

        public void OnCloseDb()
        {
            LogMessage("Отключение от базы данных");
        }

        public void BindMenu(IMenuDefinition menu)
        {
            menu.AddMenuItem("Найти дубликаты в папке", FindDuplicatesInFolder, call => true);
            LogMessage("Меню добавлено");
        }

        private void LogMessage(string message)
        {
            try
            {
                if (!string.IsNullOrEmpty(_logFilePath))
                {
                    File.AppendAllText(_logFilePath, $"[{DateTime.Now:HH:mm:ss}] {message}\n", Encoding.UTF8);
                }
            }
            catch { }
        }

        private void FindDuplicatesInFolder(INetPluginCall call)
        {
            LogMessage("=== НАЧАЛО ПОИСКА ДУБЛИКАТОВ В ПАПКЕ ===");

            StringBuilder output = new StringBuilder();
            output.AppendLine("\n" + new string('=', 50));
            output.AppendLine("=== ПОИСК ДУБЛИКАТОВ В ПАПКЕ ===");
            output.AppendLine(new string('=', 50));

            try
            {
                // 1. Получаем выделенные ID
                output.AppendLine("\n1. Получаем выделенные объекты...");
                int selectedId = GetSelectedId(call);

                if (selectedId == 0)
                {
                    output.AppendLine("ОШИБКА: Не выделена папка!");
                    SaveOutputToFile(output.ToString());
                    return;
                }

                output.AppendLine($"Выделенная папка: ID {selectedId}");

                // 2. Устанавливаем XML формат для получения данных
                output.AppendLine("\n2. Устанавливаем XML формат...");
                try
                {
                    call.RunMethod("SetFormat", new object[] { "xml" });
                    output.AppendLine("   XML формат установлен");
                }
                catch (Exception ex)
                {
                    output.AppendLine($"   Ошибка SetFormat: {ex.Message}");
                }

                // 3. Получаем информацию о папке
                output.AppendLine("\n3. Получаем информацию о папке...");
                var folderInfo = GetObjectInfo(call, selectedId, output);

                output.AppendLine($"   Название: {folderInfo.Name}");
                output.AppendLine($"   Тип: {folderInfo.Type}");
                output.AppendLine($"   Версия: {folderInfo.Version}");
                output.AppendLine($"   Состояние: {folderInfo.State}");

                // 4. Получаем словарь типов для замены ID на названия
                output.AppendLine("\n4. Получаем словарь типов...");
                _typeDictionary = GetTypeDictionary(call, output);
                output.AppendLine($"   Получено {_typeDictionary.Count} типов");

                // 5. Получаем объекты в папке
                output.AppendLine("\n5. Получаем объекты в папке...");
                List<ObjectInfo> allObjects = GetObjectsInFolder(call, selectedId, output);

                // Заменяем ID типов на названия
                foreach (var obj in allObjects)
                {
                    if (int.TryParse(obj.Type, out int typeId) && _typeDictionary.ContainsKey(typeId))
                    {
                        obj.Type = _typeDictionary[typeId];
                    }
                }

                output.AppendLine($"   Всего объектов в папке: {allObjects.Count}");

                if (allObjects.Count == 0)
                {
                    output.AppendLine("\n   Папка пуста или не удалось получить объекты.");
                    output.AppendLine("   Попробуйте:");
                    output.AppendLine("   1. Проверить, что папка содержит объекты");
                    output.AppendLine("   2. Попробовать другую папку");
                    output.AppendLine("   3. Проверить права доступа");

                    // Возвращаем бинарный формат
                    try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }

                    SaveOutputToFile(output.ToString());
                    return;
                }

                // 6. Ищем дубликаты
                output.AppendLine("\n6. Ищем дубликаты...");

                // Фильтруем объекты с названиями
                var objectsWithNames = allObjects
                    .Where(o => !string.IsNullOrEmpty(o.Name) && o.Name != $"Объект_{o.Id}")
                    .ToList();

                output.AppendLine($"   Объектов с заполненными названиями: {objectsWithNames.Count}");

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

                output.AppendLine($"   Найдено групп дубликатов:");
                output.AppendLine($"   • Критические (Название+Тип+Версия): {fullDuplicates.Count}");
                output.AppendLine($"   • Реальные дубликаты (Название+Тип): {nameTypeDuplicates.Count}");

                int totalRealDuplicates = fullDuplicates.Count + nameTypeDuplicates.Count;
                output.AppendLine($"   • ВСЕГО РЕАЛЬНЫХ ДУБЛИКАТОВ: {totalRealDuplicates}");

                // Возвращаем бинарный формат
                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }

                // 7. Сохраняем отчет
                output.AppendLine("\n7. Сохраняем отчет...");
                SaveDuplicatesReport(folderInfo, allObjects, fullDuplicates, nameTypeDuplicates);

                // 8. Выводим результаты
                output.AppendLine("\n" + new string('=', 50));
                output.AppendLine("РЕЗУЛЬТАТЫ ПОИСКА");
                output.AppendLine(new string('=', 50));
                output.AppendLine($"Папка: {folderInfo.Name}");
                output.AppendLine($"Тип: {folderInfo.Type}");
                output.AppendLine($"Всего объектов: {allObjects.Count}");
                output.AppendLine($"Групп реальных дубликатов: {totalRealDuplicates}");

                if (totalRealDuplicates > 0)
                {
                    output.AppendLine("\nНАЙДЕНЫ РЕАЛЬНЫЕ ДУБЛИКАТЫ:");

                    if (fullDuplicates.Count > 0)
                    {
                        output.AppendLine($"\nКРИТИЧЕСКИЕ ДУБЛИКАТЫ (полное совпадение):");
                        int count = 1;
                        foreach (var group in fullDuplicates.Take(3))
                        {
                            string[] parts = group.Key.Split('|');
                            output.AppendLine($"   {count++}. '{parts[0]}' (Тип: '{parts[1]}', Версия: '{parts[2]}') - {group.Count()} объектов");
                        }
                        if (fullDuplicates.Count > 3)
                            output.AppendLine($"   ... и еще {fullDuplicates.Count - 3} групп");
                    }

                    if (nameTypeDuplicates.Count > 0)
                    {
                        output.AppendLine($"\nДУБЛИКАТЫ ПО НАЗВАНИЮ И ТИПУ:");
                        int count = 1;
                        foreach (var group in nameTypeDuplicates.Take(3))
                        {
                            string[] parts = group.Key.Split('|');
                            output.AppendLine($"   {count++}. '{parts[0]}' (Тип: '{parts[1]}') - {group.Count()} объектов");
                        }
                        if (nameTypeDuplicates.Count > 3)
                            output.AppendLine($"   ... и еще {nameTypeDuplicates.Count - 3} групп");
                    }
                }
                else
                {
                    output.AppendLine("\nРЕАЛЬНЫЕ ДУБЛИКАТЫ НЕ НАЙДЕНЫ");
                }

                output.AppendLine($"\nПодробный отчет сохранен на рабочем столе");
                output.AppendLine(new string('=', 50));

                SaveOutputToFile(output.ToString());
                LogMessage("=== ПОИСК ЗАВЕРШЕН ===");

            }
            catch (Exception ex)
            {
                output.AppendLine($"\nКРИТИЧЕСКАЯ ОШИБКА: {ex.Message}");
                output.AppendLine($"StackTrace: {ex.StackTrace}");
                LogMessage($"КРИТИЧЕСКАЯ ОШИБКА: {ex.Message}");

                // Пытаемся вернуть бинарный формат в случае ошибки
                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }

                SaveOutputToFile(output.ToString());
            }
        }

        private Dictionary<int, string> GetTypeDictionary(INetPluginCall call, StringBuilder log)
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
                log.AppendLine($"   Ошибка получения словаря типов: {ex.Message}");
            }

            return typeDict;
        }

        private void SaveOutputToFile(string content)
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"duplicate_finder_output_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string filePath = Path.Combine(desktop, fileName);

                File.WriteAllText(filePath, content, Encoding.UTF8);
                LogMessage($"Вывод сохранен: {filePath}");
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка сохранения вывода: {ex.Message}");
            }
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
                LogMessage($"Ошибка получения выделенных ID: {ex.Message}");
            }

            return 0;
        }

        private ObjectInfo GetObjectInfo(INetPluginCall call, int objectId, StringBuilder log)
        {
            try
            {
                LogMessage($"Пробуем GetPropObjects для {objectId}");

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
                log.AppendLine($"   Ошибка GetPropObjects: {ex.Message}");
                LogMessage($"Ошибка GetObjectInfo: {ex.Message}");
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
                LogMessage($"Ошибка парсинга объекта: {ex.Message}");
            }

            return new ObjectInfo { Id = objectId, Name = $"Объект_{objectId}" };
        }

        private List<ObjectInfo> GetObjectsInFolder(INetPluginCall call, int folderId, StringBuilder log)
        {
            var objects = new List<ObjectInfo>();

            try
            {
                // Метод 1: GetAllLinkedObjects (рабочий метод)
                try
                {
                    log.AppendLine("   Метод 1: Пробуем GetAllLinkedObjects...");
                    object result = call.RunMethod("GetAllLinkedObjects", new object[]
                    {
                        folderId.ToString(),  // stIds
                        0                     // inParams
                    });

                    if (result is string xmlData && !string.IsNullOrEmpty(xmlData))
                    {
                        objects = ParseTreeXml(xmlData, log);
                        if (objects.Count > 0)
                        {
                            log.AppendLine($"   GetAllLinkedObjects вернул {objects.Count} объектов");
                            LogMessage($"GetAllLinkedObjects: найдено {objects.Count} объектов");
                            return objects;
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.AppendLine($"   GetAllLinkedObjects не сработал: {ex.Message}");
                }

                // Метод 2: GetTree2
                try
                {
                    log.AppendLine("   Метод 2: Пробуем GetTree2...");
                    object result = call.RunMethod("GetTree2", new object[]
                    {
                        folderId,   // inIdParent
                        0,          // inIdType
                        -1,         // inDepth
                        0           // inParams
                    });

                    if (result is string xmlData && !string.IsNullOrEmpty(xmlData))
                    {
                        objects = ParseTreeXml(xmlData, log);
                        if (objects.Count > 0)
                        {
                            log.AppendLine($"   GetTree2 вернул {objects.Count} объектов");
                            LogMessage($"GetTree2: найдено {objects.Count} объектов");
                            return objects;
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.AppendLine($"   GetTree2 не сработал: {ex.Message}");
                }

                // Метод 3: GetLinkedFast2
                try
                {
                    log.AppendLine("   Метод 3: Пробуем GetLinkedFast2...");
                    object result = call.RunMethod("GetLinkedFast2", new object[]
                    {
                        folderId.ToString(),  // stIds
                        "",                   // stLinkTypes
                        0,                    // inDirection (0 = все направления)
                        -1,                   // inDepth (-1 = все уровни)
                        0                     // inFlags
                    });

                    if (result is string xmlData && !string.IsNullOrEmpty(xmlData))
                    {
                        objects = ParseTreeXml(xmlData, log);
                        if (objects.Count > 0)
                        {
                            log.AppendLine($"   GetLinkedFast2 вернул {objects.Count} объектов");
                            LogMessage($"GetLinkedFast2: найдено {objects.Count} объектов");
                            return objects;
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.AppendLine($"   GetLinkedFast2 не сработал: {ex.Message}");
                }

                // Метод 4: GetLinkedObjects2
                try
                {
                    log.AppendLine("   Метод 4: Пробуем GetLinkedObjects2...");
                    object result = call.RunMethod("GetLinkedObjects2", new object[]
                    {
                        folderId.ToString(),  // stIds
                        1,                    // inDirection (1 = вниз)
                        -1,                   // inDepth
                        0                     // inFlags
                    });

                    if (result is string xmlData && !string.IsNullOrEmpty(xmlData))
                    {
                        objects = ParseTreeXml(xmlData, log);
                        if (objects.Count > 0)
                        {
                            log.AppendLine($"   GetLinkedObjects2 вернул {objects.Count} объектов");
                            LogMessage($"GetLinkedObjects2: найдено {objects.Count} объектов");
                            return objects;
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.AppendLine($"   GetLinkedObjects2 не сработал: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                log.AppendLine($"   Общая ошибка при получении объектов: {ex.Message}");
                LogMessage($"Ошибка GetObjectsInFolder: {ex.Message}");
            }

            return objects;
        }

        private List<ObjectInfo> ParseTreeXml(string xmlData, StringBuilder log)
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
                log.AppendLine($"   Ошибка парсинга XML: {ex.Message}");
            }

            return objects;
        }

        private void SaveDuplicatesReport(ObjectInfo folderInfo,
                                         List<ObjectInfo> allObjects,
                                         List<IGrouping<string, ObjectInfo>> fullDuplicates,
                                         List<IGrouping<string, ObjectInfo>> nameTypeDuplicates)
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string safeName = ReplaceInvalidChars(folderInfo.Name);
                string fileName = $"duplicates_report_{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string filePath = Path.Combine(desktop, fileName);

                var sb = new StringBuilder();
                sb.AppendLine(new string('=', 70));
                sb.AppendLine("ОТЧЕТ О ПОИСКЕ ДУБЛИКАТОВ В ПАПКЕ");
                sb.AppendLine(new string('=', 70));
                sb.AppendLine($"Дата: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
                sb.AppendLine($"Папка: {folderInfo.Name} (ID: {folderInfo.Id})");
                sb.AppendLine($"Тип: {folderInfo.Type}");
                sb.AppendLine($"Версия: {folderInfo.Version}");
                sb.AppendLine($"Состояние: {folderInfo.State}");
                sb.AppendLine();
                sb.AppendLine(new string('-', 70));
                sb.AppendLine();

                sb.AppendLine($"ВСЕГО ОБЪЕКТОВ В ПАПКЕ: {allObjects.Count}");
                sb.AppendLine();
                sb.AppendLine($"НАЙДЕНО ГРУПП ДУБЛИКАТОВ:");
                sb.AppendLine($"• Критические (Название+Тип+Версия): {fullDuplicates.Count}");
                sb.AppendLine($"• Реальные дубликаты (Название+Тип): {nameTypeDuplicates.Count}");
                sb.AppendLine($"• ВСЕГО РЕАЛЬНЫХ ДУБЛИКАТОВ: {fullDuplicates.Count + nameTypeDuplicates.Count}");
                sb.AppendLine();

                int totalDuplicates = fullDuplicates.Sum(g => g.Count()) +
                                    nameTypeDuplicates.Sum(g => g.Count());

                if (totalDuplicates > 0)
                {
                    sb.AppendLine($"ВСЕГО ДУБЛИРУЮЩИХ ОБЪЕКТОВ (реальные дубликаты): {totalDuplicates}");
                    sb.AppendLine();

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
                        sb.AppendLine("РЕАЛЬНЫЕ ДУБЛИКАТЫ (одинаковое название и тип):");
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
                    sb.AppendLine();
                }

                sb.AppendLine("СТАТИСТИКА ВСЕХ ОБЪЕКТОВ В ПАПКЕ:");
                sb.AppendLine(new string('-', 70));

                int count = 1;
                foreach (var obj in allObjects.OrderBy(o => o.Type).ThenBy(o => o.Name).ThenBy(o => o.Version))
                {
                    sb.AppendLine($"#{count++}");
                    sb.AppendLine($"  ID: {obj.Id}");
                    sb.AppendLine($"  Тип: {obj.Type}");
                    sb.AppendLine($"  Название: {obj.Name}");
                    sb.AppendLine($"  Версия: {obj.Version}");
                    sb.AppendLine($"  Состояние: {obj.State}");
                    sb.AppendLine(new string('-', 40));
                }

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                LogMessage($"Отчет сохранен: {filePath}");
            }
            catch (Exception ex)
            {
                LogMessage($"Ошибка сохранения отчета: {ex.Message}");
            }
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