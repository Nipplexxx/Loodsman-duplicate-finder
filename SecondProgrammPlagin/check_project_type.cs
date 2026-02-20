using Ascon.Plm.Loodsman.PluginSDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Runtime.InteropServices;

namespace DeepDuplicateFinder
{
    [LoodsmanPlugin]
    public class DeepDuplicateFinder : ILoodsmanNetPlugin
    {
        //MessageBox
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr h, string m, string c, int type);

        private Dictionary<int, string> _typeDictionary = new Dictionary<int, string>();

        // Поля для хранения ID типов
        private int _preparationTypeId = 0;
        private int _preparationLinkTypeId = 0; // ID типа связи "Изготавливается из"

        //Константы
        private const string MATERIAL_TYPE_NAME = "Материал по КД";
        private const string DETAIL_TYPE_NAME = "Деталь";
        private const string MATERIAL_MAIN_TYPE_NAME = "Материал основной";
        private const string PREPARATION_TYPE_NAME = "Заготовка";
        private const string PREPARATION_LINK_NAME = "Изготавливается из"; // Тип связи для заготовки

        public void PluginLoad()
        {
            MessageBox((IntPtr)0, "Плагин успешно загружен", "DeepDuplicateFinder", 0);
        }
        public void PluginUnload() { }
        public void OnConnectToDb(INetPluginCall call) { }
        public void OnCloseDb() { }

        public void BindMenu(IMenuDefinition menu)
        {
            menu.AddMenuItem("Найти множественные материалы (детали+заготовки)", Main, call => true);
        }

        private void Main(INetPluginCall call)
        {
            try
            {
                int selectedId = GetSelectedId(call);
                if (selectedId == 0)
                {
                    MessageBox((IntPtr)0, "Не выбрана папка для анализа.", "DeepDuplicateFinder", 0);
                    return;
                }

                try { call.RunMethod("SetFormat", new object[] { "xml" }); } catch { }

                var folderInfo = GetObjectInfo(call, selectedId);
                _typeDictionary = GetTypeDictionary(call);

                // Получаем ID всех нужных типов
                int detailTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                _preparationTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(PREPARATION_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                int materialTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(MATERIAL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                int mainMaterialTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(MATERIAL_MAIN_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;

                // ID типа связи "Изготавливается из"
                _preparationLinkTypeId = GetLinkTypeIdByName(call, PREPARATION_LINK_NAME);

                Console.WriteLine(Convert.ToString(call), $"ID детали: {detailTypeId}");
                Console.WriteLine(Convert.ToString(call), $"ID заготовки: {_preparationTypeId}");
                Console.WriteLine(Convert.ToString(call), $"ID связи 'Изготавливается из': {_preparationLinkTypeId}");

                // Получаем все объекты в папке
                List<ObjectInfo> allObjects = GetAllObjectsIncludingIndirect(call, selectedId);

                // ПОИСК ЗАГОТОВОК ЧЕРЕЗ СВЯЗЬ "Изготавливается из"
                if (_preparationLinkTypeId != 0)
                {
                    var preparationsFromLinks = FindPreparationsByLink(call, allObjects);

                    foreach (var prep in preparationsFromLinks)
                    {
                        if (!allObjects.Any(o => o.Id == prep.Id))
                        {
                            allObjects.Add(prep);
                            Console.WriteLine(Convert.ToString(call), $"Добавлена заготовка через связь: ID={prep.Id}, Name={prep.Name}");
                        }
                    }
                }

                // Заменяем ID типов на имена
                foreach (var obj in allObjects)
                {
                    if (int.TryParse(obj.Type, out int t) && _typeDictionary.ContainsKey(t))
                        obj.Type = _typeDictionary[t];
                }

                // Отделяем детали и заготовки
                var details = allObjects.Where(o => o.Type.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
                var preparations = allObjects.Where(o => o.Type.Equals(PREPARATION_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();

                Console.WriteLine(Convert.ToString(call), $"Всего объектов: {allObjects.Count}, деталей: {details.Count}, заготовок: {preparations.Count}");

                var reportItems = new List<ReportItem>();
                var allDebugInfo = new StringBuilder();

                int totalMaterialsFound = 0;
                int totalObjectsWithMultipleMaterials = 0;

                // Обрабатываем детали (ищем материалы по КД)
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
                                    })
                                    .ToList();

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

                // Обрабатываем заготовки (ищем основные материалы)
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
                                    })
                                    .ToList();

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

                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }

                // Выводим отладку
                Console.WriteLine(Convert.ToString(call), "--------------------------Отладка--------------------------");
                Console.WriteLine(Convert.ToString(call), allDebugInfo.ToString());

                // СОЗДАЕМ ОТЧЕТ
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
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = reportPath,
                        UseShellExecute = true
                    });
                }
                catch { }

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
                MessageBox((IntPtr)0, $"Ошибка: {ex.Message}", "DeepDuplicateFinder", 0);
                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
            }
        }

        // ПОИСК ЗАГОТОВОК ЧЕРЕЗ СВЯЗЬ "Изготавливается из"
        private List<ObjectInfo> FindPreparationsByLink(INetPluginCall call, List<ObjectInfo> existingObjects)
        {
            var result = new List<ObjectInfo>();

            try
            {
                // Берем все объекты, которые могут быть деталями (по ID типа)
                var potentialDetails = existingObjects.Where(o =>
                    int.TryParse(o.Type, out int t) && t == _typeDictionary.FirstOrDefault(x => x.Value == DETAIL_TYPE_NAME).Key
                ).ToList();

                Console.WriteLine($"Найдено потенциальных деталей: {potentialDetails.Count}");

                foreach (var detail in potentialDetails)
                {
                    // Ищем обратные связи (заготовки, которые ссылаются на деталь)
                    object res = call.RunMethod("GetLinkedObjectsForObjects", new object[]
                    {
                        detail.Id.ToString(), "", true
                    });

                    if (res is string xml && !string.IsNullOrEmpty(xml))
                    {
                        var linkedObjects = ParseLinkedObjectsXml(xml);

                        // Фильтруем по типу связи "Изготавливается из"
                        var preparationLinks = linkedObjects.Where(o =>
                            o.LinkTypeName?.Equals(PREPARATION_LINK_NAME, StringComparison.OrdinalIgnoreCase) == true).ToList();

                        foreach (var link in preparationLinks)
                        {
                            var prepInfo = GetObjectInfo(call, link.ParentId);
                            if (prepInfo.Id > 0 && !result.Any(r => r.Id == prepInfo.Id))
                            {
                                result.Add(prepInfo);
                                Console.WriteLine($"Найдена заготовка через связь: ID={prepInfo.Id}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка в FindPreparationsByLink: {ex.Message}");
            }

            return result;
        }

        // Получение ID типа связи по имени
        private int GetLinkTypeIdByName(INetPluginCall call, string linkTypeName)
        {
            try
            {
                object res = call.RunMethod("GetLinkTypes", new object[0]);
                if (res is string xml && !string.IsNullOrEmpty(xml))
                {
                    Console.WriteLine("XML от GetLinkTypes:");
                    Console.WriteLine(xml);

                    var doc = new XmlDocument();
                    doc.LoadXml(xml);
                    var rows = doc.SelectNodes("//row");

                    if (rows != null)
                    {
                        Console.WriteLine($"Найдено типов связей: {rows.Count}");
                        foreach (XmlNode row in rows)
                        {
                            string name = "";
                            int id = 0;
                            foreach (XmlAttribute a in row.Attributes)
                            {
                                if (a.Name == "c1" || a.Name == "_NAME") name = a.Value;
                                if (a.Name == "c0" || a.Name == "_ID") int.TryParse(a.Value, out id);
                            }
                            Console.WriteLine($"  - ID={id}, Name={name}");

                            if (name.Equals(linkTypeName, StringComparison.OrdinalIgnoreCase))
                                return id;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка GetLinkTypes: {ex.Message}");
            }
            return 0;
        }

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
                {
                    objects.Add(info);
                }

                try
                {
                    // Ищем дочерние объекты
                    object res = call.RunMethod("GetLinkedObjectsForObjects", new object[]
                    {
                        id.ToString(), "", false
                    });

                    if (res is string xml && !string.IsNullOrEmpty(xml))
                    {
                        var children = ParseLinkedObjectsXml(xml);
                        foreach (var c in children)
                        {
                            if (!visited.Contains(c.ChildId))
                            {
                                var childInfo = GetObjectInfo(call, c.ChildId);
                                if (childInfo.Id > 0 && !objects.Any(o => o.Id == childInfo.Id))
                                {
                                    objects.Add(childInfo);
                                }
                                queue.Enqueue(c.ChildId);
                            }
                        }
                    }

                    // Ищем родительские объекты
                    object resParent = call.RunMethod("GetLinkedObjectsForObjects", new object[]
                    {
                        id.ToString(), "", true
                    });

                    if (resParent is string xmlParent && !string.IsNullOrEmpty(xmlParent))
                    {
                        var parents = ParseLinkedObjectsXml(xmlParent);
                        foreach (var p in parents)
                        {
                            if (!visited.Contains(p.ParentId) && p.ParentId != id)
                            {
                                var parentInfo = GetObjectInfo(call, p.ParentId);
                                if (parentInfo.Id > 0 && !objects.Any(o => o.Id == parentInfo.Id))
                                {
                                    objects.Add(parentInfo);
                                }
                                queue.Enqueue(p.ParentId);
                            }
                        }
                    }
                }
                catch { }
            }

            return objects;
        }

        private List<ObjectInfo> GetMaterialLinks(INetPluginCall call, int versionId, int materialTypeId)
        {
            var list = new List<ObjectInfo>();
            try
            {
                object res = call.RunMethod("GetLinkedObjectsForObjects", new object[]
                {
                    versionId.ToString(), "", false
                });

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

                    foreach (var o in list)
                    {
                        if (int.TryParse(o.Type, out int tid) && _typeDictionary.ContainsKey(tid))
                        {
                            o.Type = _typeDictionary[tid];
                        }
                    }
                }
            }
            catch { }
            return list;
        }

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

                if (rows == null) return list;

                foreach (XmlNode row in rows)
                {
                    var info = new LinkedObjectInfo();
                    int tempInt;

                    foreach (XmlAttribute attr in row.Attributes)
                    {
                        string name = attr.Name;
                        string value = attr.Value;

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
                System.Diagnostics.Debug.WriteLine($"ParseLinkedObjectsXml ошибка: {ex.Message}");
            }
            return list;
        }

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
                    var rows = doc.SelectNodes("//row");

                    if (rows == null || rows.Count == 0)
                        rows = doc.SelectNodes("//ROOT/rowset/row");

                    if (rows == null) return dict;

                    foreach (XmlNode r in rows)
                    {
                        int id = 0; string name = "";
                        foreach (XmlAttribute a in r.Attributes)
                        {
                            string n = a.Name.ToUpper();
                            if (n == "C0" || n == "_ID" || n == "ID")
                                int.TryParse(a.Value, out id);
                            if (n == "C1" || n == "_NAME" || n == "NAME")
                                name = a.Value;
                        }
                        if (id > 0 && !string.IsNullOrEmpty(name) && !dict.ContainsKey(id))
                        {
                            dict[id] = name;
                        }
                    }
                }
            }
            catch { }
            return dict;
        }

        private int GetSelectedId(INetPluginCall call)
        {
            try
            {
                object res = call.RunMethod("CGetTreeSelectedIDs", new object[0]);

                if (res is string s && !string.IsNullOrEmpty(s))
                {
                    var ids = s.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    if (ids.Length > 0)
                    {
                        string firstId = ids[0].Trim();
                        if (int.TryParse(firstId, out int id) && id > 0)
                            return id;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox((IntPtr)0, $"Ошибка в GetSelectedId: {ex.Message}", "DeepDuplicateFinder", 0);
            }
            return 0;
        }

        private ObjectInfo GetObjectInfo(INetPluginCall call, int id)
        {
            var result = new ObjectInfo
            {
                Id = id,
                Name = $"Объект_{id}",
                Type = "Неизвестный",
                Version = "",
                State = ""
            };

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
            catch { }

            return result;
        }

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
                        string valueString = a.Value;
                        if (n == "C1" || n == "_TYPE") info.Type = valueString;
                        if (n == "C2" || n == "_PRODUCT") info.Name = valueString;
                        if (n == "C3" || n == "_VERSION") info.Version = valueString;
                        if (n == "C4" || n == "_STATE") info.State = valueString;
                    }
                    if (string.IsNullOrEmpty(info.Name))
                    {
                        info.Name = $"{info.Type} {info.Id}".Trim();
                    }
                    return info;
                }
            }
            catch { }
            return new ObjectInfo { Id = id, Name = $"Объект_{id}" };
        }
    }
}


//Единый класс для объекта
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

//Единый класс для группы материалов
public class MaterialGroup
{
    public string Key { get; set; }
    public string MaterialName { get; set; } = "";
    public string MaterialType { get; set; } = "";
    public string MaterialVersion { get; set; } = "";
    public int LinkCount { get; set; }
    public List<ObjectInfo> Materials { get; set; } = new List<ObjectInfo>();
}

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

// Класс для элемента отчета (может быть и деталь, и заготовка)
public class ReportItem
{
    public ObjectInfo Object { get; set; }
    public string ObjectTypeName { get; set; } = "";
    public List<ObjectInfo> Materials { get; set; } = new List<ObjectInfo>();
    public List<MaterialGroup> MaterialGroups { get; set; } = new List<MaterialGroup>();
    public int TotalMaterials { get; set; }
    public string MaterialTypeName { get; set; } = "";
}