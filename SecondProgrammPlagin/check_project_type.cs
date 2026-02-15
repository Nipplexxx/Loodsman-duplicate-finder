using Ascon.Plm.Loodsman.PluginSDK;
using DocumentFormat.OpenXml.VariantTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace DeepDuplicateFinder
{
    [LoodsmanPlugin]
    public class DeepDuplicateFinder : ILoodsmanNetPlugin
    {
        private Dictionary<int, string> _typeDictionary = new Dictionary<int, string>();
        private const string MATERIAL_TYPE_NAME = "Материал по КД";
        private const string DETAIL_TYPE_NAME = "Деталь";

        public void PluginLoad() 
        { 
        
        }
        public void PluginUnload() 
        { 

        }
        public void OnConnectToDb(INetPluginCall call) 
        { 
        
        }
        public void OnCloseDb() 
        {
        
        }

        public void BindMenu(IMenuDefinition menu)
        {
            menu.AddMenuItem("Найти дубликаты материалов в деталях (Материал по КД)", FindMaterialDuplicatesInDetails, call => true);
        }

        //Метод поиска дубликтов
        private void FindMaterialDuplicatesInDetails(INetPluginCall call)
        {
            try
            {
                int selectedId = GetSelectedId(call);

                if (selectedId == 0)
                {
                    Console.Write(Convert.ToString(call), "Не выбрана папка для анализа.");
                    return;
                }

                try 
                { 
                    call.RunMethod("SetFormat", new object[] { "xml" }); 
                } 
                catch 
                { 

                }

                var folderInfo = GetObjectInfo(call, selectedId);

                //-------Отладочная информация-------
                Console.Write(Convert.ToString(call), $"Анализ папки: {folderInfo.Name ?? "ID " + folderInfo.Id}...");

                _typeDictionary = GetTypeDictionary(call);

                //Материал
                var materialTypeName = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(MATERIAL_TYPE_NAME, StringComparison.OrdinalIgnoreCase));
                int materialTypeId = materialTypeName.Key;
                if (materialTypeId == 0)
                {
                    Console.WriteLine(Convert.ToString(call), "Тип 'Материал по КД' не найден.");
                    
                    try 
                    { 
                        call.RunMethod("SetFormat", new object[] { "" }); 
                    } 
                    catch 
                    { 
                    
                    }

                    return;
                }

                //Деталь
                var detailTypeKv = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase));
                int detailTypeId = detailTypeKv.Key;
                if (detailTypeId == 0)
                {
                    Console.WriteLine(Convert.ToString(call), "Тип 'Деталь' не найден.");
                    try 
                    { 
                        call.RunMethod("SetFormat", new object[] {" "}); 
                    } 
                    catch 
                    { 
                    
                    }
                    return;
                }

                List<ObjectInfo> allObjects = GetAllObjectsRecursive(call, selectedId);
                if (allObjects.Count == 0)
                {
                    try 
                    { 
                        call.RunMethod("SetFormat", new object[] { "" }); 
                    } 
                    catch 
                    { 
                    
                    }
                    Console.WriteLine(Convert.ToString(call), "В выбранной папке не найдено объектов.");
                    return;
                }

                //-------Отладочная информация-------
                Console.WriteLine(Convert.ToString(call), $"Найдено объектов: {allObjects.Count}. Обработка...");

                foreach (var obj in allObjects)
                {
                    if (int.TryParse(obj.Type, out int t) && _typeDictionary.ContainsKey(t))
                        obj.Type = _typeDictionary[t];
                }

                var details = allObjects.Where(o => o.Type.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();

                //-------Отладочная информация-------
                Console.WriteLine(Convert.ToString(call), $"Найдено деталей: {details.Count}.");

                if (details.Count == 0)
                {
                    try 
                    { 
                        call.RunMethod("SetFormat", new object[] { "" }); 
                    } 
                    catch 
                    { 
                    
                    }
                    return;
                }

                var detailsWithMaterialDuplicates = new List<DetailWithMaterialDuplicates>();
                var allDebugInfo = new StringBuilder();
                int totalMaterialsFound = 0;
                int totalDuplicateMaterials = 0;
                int totalDuplicateGroups = 0;

                foreach (var detail in details)
                {
                    var materialLinks = GetMaterialLinks(call, detail.Id, materialTypeId);
                    if (materialLinks.Count == 0) 
                    {
                        continue;
                    }

                    totalMaterialsFound += materialLinks.Count;

                    var detailDebug = DebugMaterialSearch(detail, materialLinks);
                    allDebugInfo.AppendLine(detailDebug);

                    var materialGroups = materialLinks
                        .GroupBy(m => $"{m.Name} | {m.Version}")
                        .Where(g => g.Count() > 1)
                        .ToList();

                    if (materialGroups.Count > 0)
                    {
                        var groupList = new List<MaterialGroup>();
                        foreach (var g in materialGroups)
                        {
                            var first = g.First();
                            var mg = new MaterialGroup
                            {
                                Key = g.Key,
                                MaterialName = first.Name,
                                MaterialVersion = first.Version,
                                LinkCount = g.Count(),
                                Materials = g.ToList()
                            };
                            groupList.Add(mg);
                            totalDuplicateMaterials += g.Count() - 1;
                            totalDuplicateGroups++;
                        }
                        detailsWithMaterialDuplicates.Add(new DetailWithMaterialDuplicates
                        {
                            Detail = detail,
                            MaterialGroups = groupList
                        });
                    }
                }

                try 
                { 
                    call.RunMethod("SetFormat", new object[] { "" }); 
                } 
                catch 
                { 
                
                }

                //-------Отладочная информация-------
                Console.WriteLine(Convert.ToString(call), "--------------------------Отладка--------------------------");
                Console.WriteLine(Convert.ToString(call), allDebugInfo.ToString());

                string reportPath = CreateMaterialDuplicatesHtmlReport(folderInfo, allObjects.Count, details.Count, totalMaterialsFound, totalDuplicateMaterials, detailsWithMaterialDuplicates);

                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = reportPath,
                        UseShellExecute = true
                    });
                }
                catch 
                { 
                
                }

                string finalMessage = $"Всего объектов: {allObjects.Count}\n" +
                    $"Всего деталей: {details.Count}\n" +
                    $"Найдено связей с материалами: {totalMaterialsFound}\n" +
                    $"Всего дубликатов: {totalDuplicateMaterials}\n" +
                    $"Групп дубликатов: {totalDuplicateGroups}\n" +
                    $"Деталей с дубликатами: {detailsWithMaterialDuplicates.Count}\n" +
                    $"Отчет: {reportPath}";

                Console.WriteLine(Convert.ToString(call), finalMessage);
            }
            catch (Exception ex)
            {
                try 
                { 
                    call.RunMethod("SetFormat", new object[] { "" }); 
                } 
                catch 
                { 
                
                }

                //-------Отладочная информация-------
                Console.WriteLine(Convert.ToString(call), $"Ошибка: {ex.Message}\n{ex.StackTrace}");
            }
        }

        //Обходит все объекты в папке
        private List<ObjectInfo> GetAllObjectsRecursive(INetPluginCall call, int rootId)
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
                if (info.Id > 0) objects.Add(info);

                try
                {
                    //Новый метод для получения дочерних объектов
                    object res = call.RunMethod("GetLinkedObjectsForObjects", new object[] 
                    {
                        id.ToString(),
                        "",
                        false
                    });

                    if (res is string xml && !string.IsNullOrEmpty(xml))
                    {
                        var children = ParseLinkedObjectsXml(xml);
                        foreach (var c in children)
                        {
                            if (!visited.Contains(c.ChildId))
                                queue.Enqueue(c.ChildId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //-------Отладочная информация-------
                    System.Diagnostics.Debug.WriteLine($"Error getting children for {id}: {ex.Message}");
                }
            }
            return objects;
        }

        private List<ObjectInfo> GetMaterialLinks(INetPluginCall call, int versionId, int materialTypeId)
        {
            var list = new List<ObjectInfo>();
            try
            {
                //GetLinkedObjectsForObjects для получения связанных объектов
                object res = call.RunMethod("GetLinkedObjectsForObjects", new object[] 
                {
                    versionId.ToString(),
                    "",
                    false
                });

                if (res is string xml && !string.IsNullOrEmpty(xml))
                {
                    System.Diagnostics.Debug.WriteLine($"GetLinkedObjectsForObjects XML for version {versionId}: {xml}");

                    //Парсинг результата
                    var linkedObjects = ParseLinkedObjectsXml(xml);

                    System.Diagnostics.Debug.WriteLine($"Всего связанных объектов: {linkedObjects.Count}");

                    //Фильтрация по типу материала
                    var materialLinks = linkedObjects.Where(o => o.TypeId == materialTypeId).ToList();

                    //-------Отладочная информация-------
                    System.Diagnostics.Debug.WriteLine($"Найдено материалов по TypeId={materialTypeId}: {materialLinks.Count}");

                    //Преобразование в ObjectInfo
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

                    //Замена ID типа на имя для удобства
                    foreach (var o in list)
                    {
                        if (int.TryParse(o.Type, out int tid) && _typeDictionary.ContainsKey(tid))
                        {
                            o.Type = _typeDictionary[tid];
                        }
                        
                        //-------Отладочная информация-------
                        System.Diagnostics.Debug.WriteLine($"  -> Материал: {o.Name} (Версия: {o.Version}), Type={o.Type}, LinkId={o.LinkId}");
                    }
                }
            }
            catch (Exception ex)
            {
                //-------Отладочная информация-------
                System.Diagnostics.Debug.WriteLine($"GetMaterialLinks error for version {versionId}: {ex.Message}");
            }
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
                
                if (rows == null || rows.Count == 0)
                {
                    rows = doc.SelectNodes("//data/row");
                }

                if (rows == null)
                {
                    System.Diagnostics.Debug.WriteLine("ParseLinkedObjectsXml: rows not found");
                    return list;
                }

                //-------Отладочная информация-------
                System.Diagnostics.Debug.WriteLine($"ParseLinkedObjectsXml: найдено {rows.Count} строк");

                foreach (XmlNode row in rows)
                {
                    var info = new LinkedObjectInfo();
                    int tempInt;

                    //Получение значения по индексам
                    foreach (XmlAttribute attr in row.Attributes)
                    {
                        string name = attr.Name; //Тут "c0", "c1", "c2"
                        string value = attr.Value;

                        //Парсинг по индексам
                        switch (name)
                        {
                            case "c0": //_ID_LINK
                                if (int.TryParse(value, out tempInt))
                                    info.LinkId = tempInt;
                                break;

                            case "c1": //_ID_PARENT
                                if (int.TryParse(value, out tempInt))
                                    info.ParentId = tempInt;
                                break;

                            case "c2": //_ID_CHILD
                                if (int.TryParse(value, out tempInt))
                                    info.ChildId = tempInt;
                                break;

                            case "c3": //_ID_LINK_TYPE
                                if (int.TryParse(value, out tempInt))
                                    info.LinkTypeId = tempInt;
                                break;

                            case "c4": //_LINK_TYPE_NAME
                                info.LinkTypeName = value;
                                break;

                            case "c5": //_ID_TYPE
                                if (int.TryParse(value, out tempInt))
                                    info.TypeId = tempInt;
                                break;

                            case "c6": //_PRODUCT
                                info.Product = value;
                                break;

                            case "c7": //_VERSION
                                info.Version = value;
                                break;

                            case "c8": //_ID_STATE
                                if (int.TryParse(value, out tempInt))
                                    info.StateId = tempInt;
                                break;

                            case "c14": //_ID_LOCK
                                if (int.TryParse(value, out tempInt))
                                    info.LockId = tempInt;
                                break;
                        }
                    }

                    //-------Отладочная информация-------
                    System.Diagnostics.Debug.WriteLine($"Распарсено: LinkId={info.LinkId}, ParentId={info.ParentId}, ChildId={info.ChildId}, TypeId={info.TypeId}, Product={info.Product}");

                    if (info.ChildId > 0)
                    {
                        list.Add(info);
                    }
                }

                System.Diagnostics.Debug.WriteLine($"ParseLinkedObjectsXml: успешно распарсено {list.Count} объектов");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ParseLinkedObjectsXml error: {ex.Message}");
            }
            return list;
        }



        //Метод для отладки поиска материалов в деталях
        private string DebugMaterialSearch(ObjectInfo parent, List<ObjectInfo> materials)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Деталь: {parent.Name} (ID: {parent.Id}, Версия: {parent.Version})");

            sb.AppendLine($"Связей с материалами: {materials.Count}");
            foreach (var m in materials)
            {
                sb.AppendLine($" - ID {m.Id}: {m.Name} (Версия: {m.Version}), Type={m.Type}, LinkId={m.LinkId}");
            }

            //Группируем по имени и версии (без ID)
            var groups = materials.GroupBy(m => new { m.Name, m.Version }).Where(g => g.Count() > 1).ToList();

            sb.AppendLine($"\nДубликаты (найдено групп: {groups.Count}):");

            foreach (var g in groups)
            {
                sb.AppendLine($"  {g.Key.Name} (Версия: {g.Key.Version}) → используется {g.Count()} раз(а)");
                foreach (var m in g)
                {
                    sb.AppendLine($"    - LinkId: {m.LinkId}, ID материала: {m.Id}");
                }
            }

            if (groups.Count == 0)
            {
                sb.AppendLine("  Дубликатов не найдено");
            }

            return sb.ToString();
        }

        //Создание дубликатов материалов в HTML
        private string CreateMaterialDuplicatesHtmlReport(ObjectInfo folderInfo, int totalObjects, int totalDetails, int totalMaterials, int totalDups, List<DetailWithMaterialDuplicates> dups)
        {
            try
            {
                var gen = new MaterialDuplicatesHtmlReportGenerator();
                return gen.CreateHtmlReport(folderInfo, totalObjects, totalDetails, totalMaterials, totalDups, dups, MATERIAL_TYPE_NAME);
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка создания отчёта: {ex.Message}");
            }
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

                    if (rows == null || rows.Count == 0) rows = doc.SelectNodes("//ROOT/rowset/row");
                    {

                    }

                    if (rows == null)
                    {
                        return dict;
                    }

                    foreach (XmlNode r in rows)
                    {
                        int id = 0; string name = "";
                        foreach (XmlAttribute a in r.Attributes)
                        {
                            string n = a.Name.ToUpper();
                            if (n == "C0" || n == "_ID" || n == "ID") int.TryParse(a.Value, out id);
                            {

                            }
                            if (n == "C1" || n == "_NAME" || n == "NAME") name = a.Value;
                            {

                            }
                        }
                        if (id > 0 && !string.IsNullOrEmpty(name) && !dict.ContainsKey(id))
                        {
                            dict[id] = name;
                        }
                    }
                }
            }
            catch 
            {
                return dict;
            }
            return dict;
        }

        private int GetSelectedId(INetPluginCall call)
        {
            try
            {
                object res = call.RunMethod("CGetTreeSelectedIDs", new object[0]);

                if (res is string s && !string.IsNullOrEmpty(s))
                {
                    System.Diagnostics.Debug.WriteLine($"CGetTreeSelectedIDs возвращает: {s}");

                    // Разбираем строку с ID (формат: "id1,id2,id3") если бы помогло
                    var ids = s.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    if (ids.Length > 0)
                    {
                        // Берем первый ID
                        string firstId = ids[0].Trim();
                        if (int.TryParse(firstId, out int id) && id > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"ID: {id}");
                            return id;
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("CGetTreeSelectedIDs Возвращает пустой или нестроковый результат.");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в GetSelectedId: {ex.Message}");
            }
            return 0;
        }

        private ObjectInfo GetObjectInfo(INetPluginCall call, int id)
        {
            try
            {
                object res = call.RunMethod("GetPropObjects", new object[] { id.ToString(), 0 });
                if (res is string xml && !string.IsNullOrEmpty(xml))
                {
                    return ParseObjectInfoXml(xml, id);
                }
            }
            catch 
            {
                
            }
            return new ObjectInfo { Id = id, Name = $"Объект_{id}" };
        }

        //Получение атрибутов для XML
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
                        if (n == "C1" || n == "_TYPE")
                        {
                            info.Type = valueString;
                        }
                        if (n == "C2" || n == "_PRODUCT")
                        {
                            info.Name = valueString;
                        }
                        if (n == "C3" || n == "_VERSION")
                        {
                            info.Version = valueString;
                        }
                        if (n == "C4" || n == "_STATE")
                        {
                            info.State = valueString;
                        }
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

    public class DetailWithMaterialDuplicates
    {
        public ObjectInfo Detail { get; set; }
        public List < MaterialGroup > MaterialGroups { get; set; }
    }

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
        public int LinkId { get; set; }           // _ID_LINK
        public int ParentId { get; set; }          // _ID_PARENT
        public int ChildId { get; set; }           // _ID_CHILD
        public int LinkTypeId { get; set; }        // _ID_LINK_TYPE
        public string LinkTypeName { get; set; } = ""; // _LINK_TYPE_NAME
        public int TypeId { get; set; }            // _ID_TYPE
        public string Product { get; set; } = "";  // _PRODUCT
        public string Version { get; set; } = "";  // _VERSION
        public int StateId { get; set; }           // _ID_STATE
        public double MinQuantity { get; set; }    // _MIN_QUANTITY
        public double MaxQuantity { get; set; }    // _MAX_QUANTITY
        public int AccessLevel { get; set; }       // _ACCESSLEVEL
        public int LockId { get; set; }            // _ID_LOCK
    }
}