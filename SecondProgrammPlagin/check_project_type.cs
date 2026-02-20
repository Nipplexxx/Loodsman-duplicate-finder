using Ascon.Plm.Loodsman.PluginSDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.IO;

namespace DeepDuplicateFinder
{
    /* Плагин для поиска деталей и заготовок, у которых более одного материала.
    При выборе папки анализирует все вложенные объекты (прямые и обратные связи),
    находит детали и заготовки, проверяет количество связанных материалов,
    формирует отчёт. */

    [LoodsmanPlugin]
    public class DeepDuplicateFinder : ILoodsmanNetPlugin
    {
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr h, string m, string c, int type);

        // Словарь для преобразования ID типа в имя (используется только для материалов)
        private Dictionary<int, string> _typeDictionary = new Dictionary<int, string>();
        private int _preparationTypeId = 0;          // ID типа "Заготовка"
        private int _preparationLinkTypeId = 0;      // ID типа связи "Заготовка для"

        // Константы с именами типов и связей (При смене БД требуется изменить)
        private const string MATERIAL_TYPE_NAME = "Материал по КД";
        private const string DETAIL_TYPE_NAME = "Деталь";
        private const string MATERIAL_MAIN_TYPE_NAME = "Материал основной";
        private const string PREPARATION_TYPE_NAME = "Заготовка";
        private const string PREPARATION_LINK_NAME = "Заготовка для";

        // Лог-файл
        private string LogPath => Path.Combine(Path.GetTempPath(), "DeepDuplicateFinder.log");

        public void PluginLoad()
        {
            Log("Плагин загружен");
            MessageBox((IntPtr)0, "Плагин успешно загружен", "DeepDuplicateFinder", 0);
        }

        public void PluginUnload() { }
        public void OnConnectToDb(INetPluginCall call) { }
        public void OnCloseDb() { }

        public void BindMenu(IMenuDefinition menu)
        {
            menu.AddMenuItem("Найти множественные материалы (детали+заготовки)", Main, call => true);
        }

        // Запись в лог-файл
        private void Log(string message)
        {
            File.AppendAllText(LogPath, $"{DateTime.Now:HH:mm:ss.fff}: {message}\n");
        }

        private void Main(INetPluginCall call)
        {
            try
            {
                // Получаем ID выделенного объекта (папки)
                int selectedId = GetSelectedId(call);
                if (selectedId == 0)
                {
                    MessageBox((IntPtr)0, "Не выбрана папка для анализа.", "DeepDuplicateFinder", 0);
                    return;
                }

                // Переключаем формат ответа на XML
                try 
                { 
                    call.RunMethod("SetFormat", new object[] { "xml" }); 
                } 
                catch 
                { 
                
                }

                // Получаем информацию о выбранной папке
                var folderInfo = GetObjectInfo(call, selectedId);

                // Загружаем словарь типов (ID -> имя)
                _typeDictionary = GetTypeDictionary(call);

                // Получаем ID интересующих нас типов
                int detailTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                _preparationTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(PREPARATION_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                int materialTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(MATERIAL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                int mainMaterialTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(MATERIAL_MAIN_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;

                Log($"ID детали: {detailTypeId}");
                Log($"ID заготовки: {_preparationTypeId}");

                // Собираем все объекты, достижимые из папки (прямые и обратные связи)
                List<ObjectInfo> allObjects = GetAllObjectsIncludingIndirect(call, selectedId);

                // Определяем ID типа связи "Заготовка для"
                _preparationLinkTypeId = FindPreparationLinkTypeId(call, allObjects);
                Log($"ID связи '{PREPARATION_LINK_NAME}': {_preparationLinkTypeId}");

                // Если ID связи известен, ищем заготовки через обратные связи
                if (_preparationLinkTypeId != 0)
                {
                    var preparationsFromLinks = FindPreparationsByLink(call, allObjects);
                    foreach (var prep in preparationsFromLinks)
                    {
                        if (!allObjects.Any(o => o.Id == prep.Id))
                        {
                            allObjects.Add(prep);
                            Log($"Добавлена заготовка через связь: ID={prep.Id}, Name={prep.Name}");
                        }
                    }
                }

                // Отделяем детали и заготовки
                var details = allObjects.Where(o => o.Type.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
                var preparations = allObjects.Where(o => o.Type.Equals(PREPARATION_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
                Log($"Всего объектов: {allObjects.Count}, деталей: {details.Count}, заготовок: {preparations.Count}");

                // Вывод отчета
                var reportItems = new List<ReportItem>();
                var allDebugInfo = new StringBuilder();
                int totalMaterialsFound = 0;
                int totalObjectsWithMultipleMaterials = 0;

                // Анализ деталей (материалы по КД)
                if (materialTypeId != 0)
                {
                    foreach (var detail in details)
                    {
                        var materialLinks = GetMaterialLinks(call, detail.Id, materialTypeId);
                        if (materialLinks.Count > 0)
                        {
                            totalMaterialsFound += materialLinks.Count;
                            allDebugInfo.AppendLine(DebugMaterialSearch(detail, materialLinks, "Деталь", MATERIAL_TYPE_NAME));
                            if (materialLinks.Count > 1)
                            {
                                var materialGroups = materialLinks
                                    .GroupBy(m => $"{m.Name} | {m.Version}")
                                    .Select(g => new MaterialGroup
                                    {
                                        Key = g.Key,
                                        MaterialName = g.First().Name,
                                        MaterialVersion = g.First().Version,
                                        LinkCount = g.Count(),
                                        Materials = g.ToList()
                                    }).ToList();
                                reportItems.Add(new ReportItem
                                {
                                    Object = detail,
                                    ObjectTypeName = "Деталь",
                                    Materials = materialLinks,
                                    MaterialGroups = materialGroups,
                                    TotalMaterials = materialLinks.Count,
                                    MaterialTypeName = MATERIAL_TYPE_NAME
                                });
                                totalObjectsWithMultipleMaterials++;
                            }
                        }
                    }
                }

                // Анализ заготовок (основные материалы)
                if (_preparationTypeId != 0 && mainMaterialTypeId != 0)
                {
                    foreach (var preparation in preparations)
                    {
                        var materialLinks = GetMaterialLinks(call, preparation.Id, mainMaterialTypeId);
                        if (materialLinks.Count > 0)
                        {
                            totalMaterialsFound += materialLinks.Count;
                            allDebugInfo.AppendLine(DebugMaterialSearch(preparation, materialLinks, "Заготовка", MATERIAL_MAIN_TYPE_NAME));
                            if (materialLinks.Count > 1)
                            {
                                var materialGroups = materialLinks
                                    .GroupBy(m => $"{m.Name} | {m.Version}")
                                    .Select(g => new MaterialGroup
                                    {
                                        Key = g.Key,
                                        MaterialName = g.First().Name,
                                        MaterialVersion = g.First().Version,
                                        LinkCount = g.Count(),
                                        Materials = g.ToList()
                                    }).ToList();
                                reportItems.Add(new ReportItem
                                {
                                    Object = preparation,
                                    ObjectTypeName = "Заготовка",
                                    Materials = materialLinks,
                                    MaterialGroups = materialGroups,
                                    TotalMaterials = materialLinks.Count,
                                    MaterialTypeName = MATERIAL_MAIN_TYPE_NAME
                                });
                                totalObjectsWithMultipleMaterials++;
                            }
                        }
                    }
                }

                try 
                { 
                    call.RunMethod("SetFormat", new object[] { "" }); 
                } 
                catch 
                { 
                
                }

                // Отладочная информация
                Log("--------------------------Отладка--------------------------");
                Log(allDebugInfo.ToString());

                // Генерируем отчёт
                var reportGenerator = new UnifiedReportGenerator();
                string reportPath = reportGenerator.CreateReport(
                    folderInfo,
                    allObjects.Count,
                    details.Count,
                    preparations.Count,
                    totalMaterialsFound,
                    totalObjectsWithMultipleMaterials,
                    reportItems);

                try 
                { 
                    System.Diagnostics.Process.Start(reportPath); 
                } 
                catch 
                { 
                
                }

                string finalMessage =
                    $"Всего объектов: {allObjects.Count}\n" +
                    $"Деталей: {details.Count}\n" +
                    $"Заготовок: {preparations.Count}\n" +
                    $"Всего связей с материалами: {totalMaterialsFound}\n" +
                    $"Объектов с несколькими материалами: {totalObjectsWithMultipleMaterials}\n" +
                    $"Отчет: {reportPath}";
                MessageBox((IntPtr)0, finalMessage, "DeepDuplicateFinder", 0);
            }
            catch (Exception ex)
            {
                Log($"Ошибка: {ex}");
                MessageBox((IntPtr)0, $"Ошибка: {ex.Message}", "DeepDuplicateFinder", 0);
                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
            }
        }

        // Определяет ID типа связи "Заготовка для", анализируя обратные связи деталей. Используется, если ID ещё не известен
        private int FindPreparationLinkTypeId(INetPluginCall call, List<ObjectInfo> objects)
        {
            var potentialDetails = objects.Where(o => o.Type.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
            Log($"Поиск ID связи '{PREPARATION_LINK_NAME}': найдено деталей {potentialDetails.Count}");

            foreach (var detail in potentialDetails)
            {
                try
                {
                    object res = call.RunMethod("GetLinkedObjectsForObjects", new object[] { detail.Id.ToString(), "", true });
                    if (res is string xml && !string.IsNullOrEmpty(xml))
                    {
                        var linked = ParseLinkedObjectsXml(xml);
                        var prepLink = linked.FirstOrDefault(l => l.LinkTypeName?.Trim().Equals(PREPARATION_LINK_NAME, StringComparison.OrdinalIgnoreCase) == true);
                        if (prepLink != null)
                        {
                            Log($"Найден ID связи '{PREPARATION_LINK_NAME}': {prepLink.LinkTypeId} (по объекту {detail.Id})");
                            return prepLink.LinkTypeId;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Ошибка при поиске связи для объекта {detail.Id}: {ex.Message}");
                }
            }
            Log($"Не удалось найти ID связи '{PREPARATION_LINK_NAME}'");
            return 0;
        }

        // Ищет заготовки, связанные с деталями через обратную связь "Заготовка для"
        private List<ObjectInfo> FindPreparationsByLink(INetPluginCall call, List<ObjectInfo> existingObjects)
        {
            var result = new List<ObjectInfo>();
            try
            {
                Log($"Всего объектов в allObjects: {existingObjects.Count}");
                foreach (var obj in existingObjects)
                {
                    Log($" Object ID={obj.Id}, Type='{obj.Type}'");
                }

                var potentialDetails = existingObjects.Where(o => o.Type.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
                Log($"Найдено потенциальных деталей: {potentialDetails.Count}");

                foreach (var detail in potentialDetails)
                {
                    object res = call.RunMethod("GetLinkedObjectsForObjects", new object[] { detail.Id.ToString(), "", true });
                    if (res is string xml && !string.IsNullOrEmpty(xml))
                    {
                        var linkedObjects = ParseLinkedObjectsXml(xml);
                        foreach (var link in linkedObjects)
                        {
                            Log($" Связь: LinkId={link.LinkId}, ParentId={link.ParentId}, ChildId={link.ChildId}, LinkTypeId={link.LinkTypeId}, LinkTypeName={link.LinkTypeName}");
                        }

                        var preparationLinks = linkedObjects.Where(o => o.LinkTypeId == _preparationLinkTypeId).ToList();
                        foreach (var link in preparationLinks)
                        {
                            // Получаем информацию об обоих концах связи
                            var parentInfo = GetObjectInfo(call, link.ParentId);
                            var childInfo = GetObjectInfo(call, link.ChildId);

                            int prepId = 0;
                            if (parentInfo.Type.Equals(PREPARATION_TYPE_NAME, StringComparison.OrdinalIgnoreCase))
                                prepId = link.ParentId;
                            else if (childInfo.Type.Equals(PREPARATION_TYPE_NAME, StringComparison.OrdinalIgnoreCase))
                                prepId = link.ChildId;
                            else
                            {
                                Log($" Не удалось определить заготовку: ParentType={parentInfo.Type}, ChildType={childInfo.Type}");
                                continue;
                            }

                            var prepInfo = GetObjectInfo(call, prepId);
                            if (prepInfo.Id > 0 && !result.Any(r => r.Id == prepInfo.Id))
                            {
                                result.Add(prepInfo);
                                Log($"Найдена заготовка: ID={prepInfo.Id}, Name={prepInfo.Name}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка в FindPreparationsByLink: {ex.Message}");
            }
            return result;
        }

        // Рекурсивный обход графа связей, начиная с корневого объекта. Собирает все достижимые объекты (прямые и обратные связи)
        private List<ObjectInfo> GetAllObjectsIncludingIndirect(INetPluginCall call, int rootId)
        {
            var objects = new List<ObjectInfo>();
            var visited = new HashSet<int>();
            var queue = new Queue<int>();
            queue.Enqueue(rootId);

            while (queue.Count > 0)
            {
                int id = queue.Dequeue();
                if (visited.Contains(id)) continue;
                visited.Add(id);

                var info = GetObjectInfo(call, id);
                if (info.Id > 0 && !objects.Any(o => o.Id == info.Id))
                    objects.Add(info);

                try
                {
                    // Прямые связи (дочерние объекты)
                    object res = call.RunMethod("GetLinkedObjectsForObjects", new object[] { id.ToString(), "", false });
                    if (res is string xml && !string.IsNullOrEmpty(xml))
                    {
                        var children = ParseLinkedObjectsXml(xml);
                        foreach (var c in children)
                        {
                            if (!visited.Contains(c.ChildId))
                                queue.Enqueue(c.ChildId);
                        }
                    }

                    // Обратные связи (родительские объекты)
                    object resParent = call.RunMethod("GetLinkedObjectsForObjects", new object[] { id.ToString(), "", true });
                    if (resParent is string xmlParent && !string.IsNullOrEmpty(xmlParent))
                    {
                        var parents = ParseLinkedObjectsXml(xmlParent);
                        foreach (var p in parents)
                        {
                            if (!visited.Contains(p.ParentId) && p.ParentId != id)
                                queue.Enqueue(p.ParentId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Ошибка при обходе объекта {id}: {ex.Message}");
                }
            }
            return objects;
        }

        // Получает список материалов, связанных с объектом (деталью или заготовкой) через прямые связи
        private List<ObjectInfo> GetMaterialLinks(INetPluginCall call, int versionId, int materialTypeId)
        {
            var list = new List<ObjectInfo>();
            try
            {
                object res = call.RunMethod("GetLinkedObjectsForObjects", new object[] { versionId.ToString(), "", false });
                if (res is string xml && !string.IsNullOrEmpty(xml))
                {
                    var linkedObjects = ParseLinkedObjectsXml(xml);
                    var materialLinks = linkedObjects.Where(o => o.TypeId == materialTypeId).ToList();

                    list = materialLinks.Select(o => new ObjectInfo
                    {
                        Id = o.ChildId,
                        Name = o.Product,
                        Type = o.TypeId.ToString(),
                        Version = o.Version,
                        State = o.StateId.ToString(),
                        LinkId = o.LinkId,
                        ParentId = o.ParentId
                    }).ToList();

                    // Заменяем ID типа на имя
                    foreach (var o in list)
                    {
                        if (int.TryParse(o.Type, out int tid) && _typeDictionary.ContainsKey(tid))
                            o.Type = _typeDictionary[tid];
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка в GetMaterialLinks для {versionId}: {ex.Message}");
            }
            return list;
        }

        // Парсит XML
        private List<LinkedObjectInfo> ParseLinkedObjectsXml(string xmlData)
        {
            var list = new List<LinkedObjectInfo>();
            try
            {
                var doc = new XmlDocument();
                doc.LoadXml(xmlData);
                XmlNodeList rows = doc.SelectNodes("//row");
                if (rows == null || rows.Count == 0)
                {
                    rows = doc.SelectNodes("//ROOT/rowset/row");
                }

                if (rows == null)
                {
                    return list;
                }

                foreach (XmlNode row in rows)
                {
                    var info = new LinkedObjectInfo();
                    int tempInt;
                    foreach (XmlAttribute attr in row.Attributes)
                    {
                        string name = attr.Name;
                        string value = attr.Value.Trim();
                        switch (name)
                        {
                            case "c0": if (int.TryParse(value, out tempInt)) info.LinkId = tempInt; break;
                            case "c1": if (int.TryParse(value, out tempInt)) info.ParentId = tempInt; break;
                            case "c2": if (int.TryParse(value, out tempInt)) info.ChildId = tempInt; break;
                            case "c3": if (int.TryParse(value, out tempInt)) info.LinkTypeId = tempInt; break;
                            case "c4": info.LinkTypeName = value; break;
                            case "c5": if (int.TryParse(value, out tempInt)) info.TypeId = tempInt; break;
                            case "c6": info.Product = value; break;
                            case "c7": info.Version = value; break;
                            case "c8": if (int.TryParse(value, out tempInt)) info.StateId = tempInt; break;
                            case "c14": if (int.TryParse(value, out tempInt)) info.LockId = tempInt; break;
                        }
                    }
                    if (info.ChildId > 0)
                    {
                        list.Add(info);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"ParseLinkedObjectsXml ошибка: {ex.Message}");
            }
            return list;
        }

        // Формирует отладочную строку для объекта и его материалов
        private string DebugMaterialSearch(ObjectInfo obj, List<ObjectInfo> materials, string objTypeName, string materialTypeName)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{objTypeName}: {obj.Name} (ID: {obj.Id}, Версия: {obj.Version})");
            sb.AppendLine($"Тип материала: {materialTypeName}");
            sb.AppendLine($"Связей: {materials.Count}");
            foreach (var m in materials)
            {
                sb.AppendLine($" - ID {m.Id}: {m.Name} (Версия: {m.Version})");
            }

            var groups = materials.GroupBy(m => new { m.Name, m.Version }).Where(g => g.Count() > 1).ToList();
            if (groups.Count > 0)
            {
                sb.AppendLine($"\nДубликатов групп: {groups.Count}");
                foreach (var g in groups)
                {
                    sb.AppendLine($"  {g.Key.Name} (Версия: {g.Key.Version}) → {g.Count()} раз");
                }
            }
            return sb.ToString();
        }

        // Загружает словарь типов
        private Dictionary<int, string> GetTypeDictionary(INetPluginCall call)
        {
            var dict = new Dictionary<int, string>();
            try
            {
                object res = call.RunMethod("GetTypeList", new object[0]);
                if (res is string xml && !string.IsNullOrEmpty(xml))
                {
                    var doc = new XmlDocument();
                    doc.LoadXml(xml);
                    var rows = doc.SelectNodes("//row") ?? doc.SelectNodes("//ROOT/rowset/row");
                    if (rows != null)
                    {
                        foreach (XmlNode r in rows)
                        {
                            int id = 0; string name = "";
                            foreach (XmlAttribute a in r.Attributes)
                            {
                                string n = a.Name.ToUpper();
                                string val = a.Value.Trim();
                                if (n == "C0" || n == "_ID" || n == "ID") int.TryParse(val, out id);
                                if (n == "C1" || n == "_NAME" || n == "NAME") name = val;
                            }
                            if (id > 0 && !string.IsNullOrEmpty(name) && !dict.ContainsKey(id))
                            {
                                dict[id] = name;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"GetTypeDictionary ошибка: {ex.Message}");
            }
            return dict;
        }

        // Возвращает ID первого выделенного в дереве объекта
        private int GetSelectedId(INetPluginCall call)
        {
            try
            {
                object res = call.RunMethod("CGetTreeSelectedIDs", new object[0]);
                if (res is string s && !string.IsNullOrEmpty(s))
                {
                    var ids = s.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    if (ids.Length > 0 && int.TryParse(ids[0].Trim(), out int id) && id > 0)
                    {
                        return id;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"GetSelectedId ошибка: {ex.Message}");
            }
            return 0;
        }

        // Получает базовую информацию об объекте по его ID
        private ObjectInfo GetObjectInfo(INetPluginCall call, int id)
        {
            var result = new ObjectInfo { Id = id, Name = $"Объект_{id}", Type = "Неизвестный", Version = "", State = "" };
            try
            {
                object res = call.RunMethod("GetPropObjects", new object[] { id.ToString(), 0 });
                if (res is string xml && !string.IsNullOrEmpty(xml))
                {
                    var parsed = ParseObjectInfoXml(xml, id);
                    if (parsed.Id > 0)
                    {
                        result.Id = parsed.Id;
                        result.Name = parsed.Name;
                        result.Type = parsed.Type;
                        result.Version = parsed.Version;
                        result.State = parsed.State;
                    }
                }
            }
            catch 
            { 
                
            }
            return result;
        }

        // Парсит XML от GetPropObjects
        private ObjectInfo ParseObjectInfoXml(string xml, int id)
        {
            try
            {
                var doc = new XmlDocument();
                doc.LoadXml(xml);
                var row = doc.SelectSingleNode("//row") ?? doc.SelectSingleNode("//ROOT/rowset/row");
                if (row != null)
                {
                    var info = new ObjectInfo { Id = id };
                    foreach (XmlAttribute a in row.Attributes)
                    {
                        string n = a.Name.ToUpper();
                        string val = a.Value.Trim();
                        if (n == "C1" || n == "_TYPE") info.Type = val;
                        if (n == "C2" || n == "_PRODUCT") info.Name = val;
                        if (n == "C3" || n == "_VERSION") info.Version = val;
                        if (n == "C4" || n == "_STATE") info.State = val;
                    }
                    if (string.IsNullOrEmpty(info.Name))
                    {
                        info.Name = $"{info.Type} {info.Id}".Trim();
                    }
                    return info;
                }
            }
            catch 
            { 
                
            }
            return new ObjectInfo { Id = id, Name = $"Объект_{id}" };
        }
    }

    // Базовая информация об объекте
    public class ObjectInfo
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string Version { get; set; } = "";
        public string State { get; set; } = "";
        public int LinkId { get; set; }
        public int ParentId { get; set; }
    }

    // Группа одинаковых материалов (для отчёта)
    public class MaterialGroup
    {
        public string Key { get; set; }
        public string MaterialName { get; set; } = "";
        public string MaterialType { get; set; } = "";
        public string MaterialVersion { get; set; } = "";
        public int LinkCount { get; set; }
        public List<ObjectInfo> Materials { get; set; } = new List<ObjectInfo>();
    }

    // Информация о связи, возвращаемая GetLinkedObjectsForObjects
    public class LinkedObjectInfo
    {
        public int LinkId { get; set; }
        public int ParentId { get; set; }
        public int ChildId { get; set; }
        public int LinkTypeId { get; set; }
        public string LinkTypeName { get; set; } = "";
        public int TypeId { get; set; }
        public string Product { get; set; } = "";
        public string Version { get; set; } = "";
        public int StateId { get; set; }
        public double MinQuantity { get; set; }
        public double MaxQuantity { get; set; }
        public int AccessLevel { get; set; }
        public int LockId { get; set; }
    }

    // Элемент отчёта (деталь или заготовка с её материалами)
    public class ReportItem
    {
        public ObjectInfo Object { get; set; }
        public string ObjectTypeName { get; set; } = "";
        public List<ObjectInfo> Materials { get; set; } = new List<ObjectInfo>();
        public List<MaterialGroup> MaterialGroups { get; set; } = new List<MaterialGroup>();
        public int TotalMaterials { get; set; }
        public string MaterialTypeName { get; set; } = "";
    }
}