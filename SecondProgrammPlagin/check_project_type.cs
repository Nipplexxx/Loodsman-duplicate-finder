using Ascon.Plm.Loodsman.PluginSDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Runtime.InteropServices;
using System.Windows;

namespace DeepDuplicateFinder
{
    [LoodsmanPlugin]
    public class DeepDuplicateFinder : ILoodsmanNetPlugin
    {
        //Для MessageBox, но круче
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr h, string m, string c, int type);


        private Dictionary<int, string> _typeDictionary = new Dictionary<int, string>();

        //Костанты
        private const string MATERIAL_TYPE_NAME = "Материал по КД";
        private const string DETAIL_TYPE_NAME = "Деталь";
        private const string MATERIAL_MAIN_TYPE_NAME = "Материал основной";
        private const string PREPARATION_TYPE_NAME = "Заготовки";

        public void PluginLoad() 
        {
            MessageBox((IntPtr)0, "Плагин успешно загружен", "DeepDuplicateFinder", 0);
        }
        public void PluginUnload() { }
        public void OnConnectToDb(INetPluginCall call) { }
        public void OnCloseDb() { }

        public void BindMenu(IMenuDefinition menu)
        {
            menu.AddMenuItem("Найти дубликаты материалов в деталях и заготовках (АВО)", FindMaterialDuplicatesInDetails, call => true);
        }

        //Метод поиска деталей с несколькими материалами
        private void FindMaterialDuplicatesInDetails(INetPluginCall call)
        {
            try
            {
                int selectedId = GetSelectedId(call);

                if (selectedId == 0)
                {
                    MessageBox((IntPtr)0, $"{Convert.ToString(call)} \n Не выбрана папка для анализа.", "DeepDuplicateFinder", 0);
                    return;
                }

                try
                {
                    call.RunMethod("SetFormat", new object[] { "xml" });
                }
                catch { }

                var folderInfo = GetObjectInfo(call, selectedId);

                //Отладка в консоле
                Console.WriteLine(Convert.ToString(call), $"Анализ папки: {folderInfo.Name} (ID: {folderInfo.Id})");

                _typeDictionary = GetTypeDictionary(call);

                //Деталь
                var detailTypeKv = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase));
                int detailTypeId = detailTypeKv.Key;
                if (detailTypeId == 0)
                {
                    MessageBox((IntPtr)0, $"{Convert.ToString(call)} \n Тип 'Деталь' не найден.", "DeepDuplicateFinder", 0);
                    try 
                    { 
                        call.RunMethod("SetFormat", new object[] { " " }); 
                    } 
                    catch 
                    { 

                    }
                    return;
                }

                //Материал
                var materialTypeName = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(MATERIAL_TYPE_NAME, StringComparison.OrdinalIgnoreCase));
                int materialTypeId = materialTypeName.Key;
                if (materialTypeId == 0)
                {
                    MessageBox((IntPtr)0, $"{Convert.ToString(call)} \n Тип 'Материал по КД' не найден.", "DeepDuplicateFinder", 0);
                    try 
                    { 
                        call.RunMethod("SetFormat", new object[] { "" }); 
                    } 
                    catch 
                    { 
                    
                    }
                    return;
                }

                List<ObjectInfo> allObjects = GetAllObjectsRecursive(call, selectedId);
                if (allObjects.Count == 0)
                {
                    MessageBox((IntPtr)0, $"{Convert.ToString(call)} \n В папке не найдено объектов.", "DeepDuplicateFinder", 0);
                    try 
                    { 
                        call.RunMethod("SetFormat", new object[] { "" }); 
                    } 
                    catch 
                    { 
                    
                    }
                    return;
                }

                //Заменяем ID типов на имена
                foreach (var obj in allObjects)
                {
                    if (int.TryParse(obj.Type, out int t) && _typeDictionary.ContainsKey(t))
                    {
                        obj.Type = _typeDictionary[t];
                    }
                }

                var details = allObjects.Where(o => o.Type.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
                Console.WriteLine(Convert.ToString(call), $"Найдено деталей: {details.Count}");

                if (details.Count == 0)
                {
                    MessageBox((IntPtr)0, $"{Convert.ToString(call)} \n Детали не найдены.", "DeepDuplicateFinder", 0);
                    try 
                    { 
                        call.RunMethod("SetFormat", new object[] { "" }); 
                    } 
                    catch 
                    { 
                        
                    }
                    return;
                }

                //Список деталей с проблемами (материалов > 1)
                var detailsWithMultipleMaterials = new List<DetailWithMultipleMaterials>();
                var allDebugInfo = new StringBuilder();
                int totalMaterialsFound = 0;
                int totalDetailsWithMultipleMaterials = 0;

                foreach (var detail in details)
                {
                    var materialLinks = GetMaterialLinks(call, detail.Id, materialTypeId);

                    if (materialLinks.Count > 0)
                    {
                        totalMaterialsFound += materialLinks.Count;

                        var detailDebug = DebugMaterialSearch(detail, materialLinks);
                        allDebugInfo.AppendLine(detailDebug);

                        //Если материалов больше 1 - добавляем в список проблемных
                        if (materialLinks.Count > 1)
                        {
                            //Группируем материалы для удобства отображения
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

                            detailsWithMultipleMaterials.Add(new DetailWithMultipleMaterials
                            {
                                Detail = detail,
                                Materials = materialLinks,
                                MaterialGroups = materialGroups,
                                TotalMaterials = materialLinks.Count
                            });

                            totalDetailsWithMultipleMaterials++;
                        }
                    }
                }

                try
                {
                    call.RunMethod("SetFormat", new object[] { "" });
                }
                catch { }

                //--------------------------Отладка--------------------------
                //Выводим отладочную информацию
                Console.WriteLine(Convert.ToString(call), "--------------------------Отладка--------------------------");
                Console.WriteLine(Convert.ToString(call), allDebugInfo.ToString());

                //Создание отчета - ВЫЗОВ МЕТОДА (не объявление!)
                string reportPath = CreateMaterialReport(
                    folderInfo,
                    allObjects.Count,
                    details.Count,
                    totalMaterialsFound,
                    detailsWithMultipleMaterials);

                //Открываем отчет
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = reportPath,
                        UseShellExecute = true
                    });
                }
                catch { }

                string finalMessage = $"Всего объектов: {allObjects.Count}\n" + $"Всего деталей: {details.Count}\n" + $"Найдено связей с материалами: {totalMaterialsFound}\n" + $"Деталей с несколькими материалами (>1): {totalDetailsWithMultipleMaterials}\n" + $"Отчет: {reportPath}";
                
                MessageBox((IntPtr)0, $"{Convert.ToString(call)} \n Ошибка: {finalMessage}", "DeepDuplicateFinder", 0);
            }
            catch (Exception ex)
            {
                MessageBox((IntPtr)0, $"{Convert.ToString(call)} \n Ошибка: {ex.Message}", "DeepDuplicateFinder", 0);
                try 
                { 
                    call.RunMethod("SetFormat", new object[] { "" });
                } 
                catch 
                { 
                    
                }
            }
        }

        //Создание отчета
        private string CreateMaterialReport(
            ObjectInfo folderInfo,
            int totalObjects,
            int totalDetails,
            int totalMaterials,
            List<DetailWithMultipleMaterials> problemDetails)
        {
            try
            {
                var gen = new MaterialDuplicatesHtmlReportGenerator();

                // Преобразуем List<DetailWithMultipleMaterials> в List<DetailWithMaterialDuplicates>
                var oldFormatList = problemDetails.Select(p => new DetailWithMaterialDuplicates
                {
                    Detail = p.Detail,
                    MaterialGroups = p.MaterialGroups
                }).ToList();

                return gen.CreateHtmlReport(folderInfo, totalObjects, totalDetails, totalMaterials, 0, oldFormatList, MATERIAL_TYPE_NAME);
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка создания отчёта: {ex.Message}");
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
                        id.ToString(), "", false
                    });

                    if (res is string xml && !string.IsNullOrEmpty(xml))
                    {
                        var children = ParseLinkedObjectsXml(xml);
                        foreach (var c in children)
                        {
                            if (!visited.Contains(c.ChildId))
                            {
                                queue.Enqueue(c.ChildId);
                            }
                        }
                    }
                }
                catch (Exception ex) { }
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

                    System.Diagnostics.Debug.WriteLine($"Всего связанных объектов: {linkedObjects.Count}");

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

                if (rows == null)
                {
                    MessageBox((IntPtr)0, "ParseLinkedObjectsXml выдал ошибку: \n строки не найдены", "DeepDuplicateFinder", 0);
                    return list;
                }

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
                MessageBox((IntPtr)0, $"ParseLinkedObjectsXml выдал ошибку: \n {ex.Message}", "DeepDuplicateFinder", 0);
            }
            return list;
        }

        //Метод для отладки поиска материалов в деталях (удалить)
        private string DebugMaterialSearch(ObjectInfo parent, List<ObjectInfo> materials)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Деталь: {parent.Name} (ID: {parent.Id}, Версия: {parent.Version})");

            sb.AppendLine($"Связей с материалами: {materials.Count}");
            foreach (var m in materials)
            {
                sb.AppendLine($" - ID {m.Id}: {m.Name} (Версия: {m.Version}), Type={m.Type}, LinkId={m.LinkId}");
            }

            var groups = materials.GroupBy(m => new { m.Name, m.Version }).Where(g => g.Count() > 1).ToList();

            sb.AppendLine($"\n Дубликаты (найдено групп: {groups.Count}):");

            foreach (var g in groups)
            {
                sb.AppendLine($"  {g.Key.Name} (Версия: {g.Key.Version}) → используется {g.Count()} раз(а)");
                foreach (var m in g)
                {
                    sb.AppendLine($"LinkId: {m.LinkId}, ID материала: {m.Id}");
                }
            }

            if (groups.Count == 0)
            {
                sb.AppendLine("Дубликатов не найдено");
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
            catch 
            { 
            
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
                    var ids = s.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    if (ids.Length > 0)
                    {
                        string firstId = ids[0].Trim();
                        if (int.TryParse(firstId, out int id) && id > 0)
                        {
                            return id;
                        }
                    }
                }
                else
                {
                    MessageBox((IntPtr)0, $"CGetTreeSelectedIDs выдал ошибку: \n озвращает пустой или нестроковый результат.", "DeepDuplicateFinder", 0);
                }
            }
            catch (Exception ex)
            {
                MessageBox((IntPtr)0, $"Ошибка в GetSelectedId: \n {ex.Message}", "DeepDuplicateFinder", 0);
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
            catch { }
            return new ObjectInfo { Id = id, Name = $"Объект_{id}" };
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

    //Класс для деталей с несколькими материалами
    public class DetailWithMultipleMaterials
    {
        public ObjectInfo Detail { get; set; }
        public List<ObjectInfo> Materials { get; set; } = new List<ObjectInfo>();
        public List<MaterialGroup> MaterialGroups { get; set; } = new List<MaterialGroup>();
        public int TotalMaterials { get; set; }
    }

    public class DetailWithMaterialDuplicates
    {
        public ObjectInfo Detail { get; set; }
        public List<MaterialGroup> MaterialGroups { get; set; }
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