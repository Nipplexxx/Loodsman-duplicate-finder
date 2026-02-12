using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Ascon.Plm.Loodsman.PluginSDK;
using System.IO;
using System.Net;

namespace DeepDuplicateFinder
{
    [LoodsmanPlugin]
    public class DeepDuplicateFinder : ILoodsmanNetPlugin
    {
        private Dictionary<int, string> _typeDictionary = new Dictionary<int, string>();
        private const string MATERIAL_TYPE_NAME = "Материал по КД";
        private const string DETAIL_TYPE_NAME = "Деталь";

        public void PluginLoad() { }
        public void PluginUnload() { }
        public void OnConnectToDb(INetPluginCall call) { }
        public void OnCloseDb() { }

        public void BindMenu(IMenuDefinition menu)
        {
            menu.AddMenuItem("Найти дубликаты материалов в деталях", FindMaterialDuplicatesInDetails, call => true);
        }

        private void FindMaterialDuplicatesInDetails(INetPluginCall call)
        {
            try
            {
                int selectedId = GetSelectedId(call);
                if (selectedId == 0)
                {
                    ShowMessage(call, "Не выбрана папка для анализа.");
                    return;
                }

                try { call.RunMethod("SetFormat", new object[] { "xml" }); } catch { }

                var folderInfo = GetObjectInfo(call, selectedId);
                ShowMessage(call, $"Анализ папки: {folderInfo.Name ?? "ID " + folderInfo.Id}...");

                _typeDictionary = GetTypeDictionary(call);

                var materialTypeKv = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(MATERIAL_TYPE_NAME, StringComparison.OrdinalIgnoreCase));
                int materialTypeId = materialTypeKv.Key;
                if (materialTypeId == 0)
                {
                    ShowMessage(call, "Тип 'Материал по КД' не найден.");
                    try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
                    return;
                }

                var detailTypeKv = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase));
                int detailTypeId = detailTypeKv.Key;
                if (detailTypeId == 0)
                {
                    ShowMessage(call, "Тип 'Деталь' не найден.");
                    try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
                    return;
                }

                List<ObjectInfo> allObjects = GetAllObjectsRecursive(call, selectedId);
                if (allObjects.Count == 0)
                {
                    try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
                    ShowMessage(call, "В выбранной папке не найдено объектов.");
                    return;
                }

                ShowMessage(call, $"Найдено объектов: {allObjects.Count}. Обработка...");

                foreach (var obj in allObjects)
                {
                    if (int.TryParse(obj.Type, out int t) && _typeDictionary.ContainsKey(t))
                        obj.Type = _typeDictionary[t];
                }

                var details = allObjects.Where(o => o.Type.Equals(DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
                ShowMessage(call, $"Найдено деталей: {details.Count}.");

                if (details.Count == 0)
                {
                    try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
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
                    if (materialLinks.Count == 0) continue;

                    totalMaterialsFound += materialLinks.Count;

                    var detailDebug = DebugMaterialSearch(detail, materialLinks);
                    allDebugInfo.AppendLine(detailDebug);

                    var materialGroups = materialLinks
                        .GroupBy(m => m.Name.Trim() + " | " + m.Version.Trim())
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
                                MaterialType = first.Type,
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

                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }

                ShowMessage(call, "=== ОТЛАДОЧНАЯ ИНФОРМАЦИЯ ===");
                ShowMessage(call, allDebugInfo.ToString());

                string reportPath = CreateMaterialDuplicatesHtmlReport(
                    folderInfo, allObjects.Count, details.Count,
                    totalMaterialsFound, totalDuplicateMaterials, detailsWithMaterialDuplicates);

                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = reportPath,
                        UseShellExecute = true
                    });
                }
                catch { }

                string finalMessage = $"Анализ завершен.\n" +
                    $"Всего объектов: {allObjects.Count}\n" +
                    $"Всего деталей: {details.Count}\n" +
                    $"Найдено связей с материалами: {totalMaterialsFound}\n" +
                    $"Всего дубликатов: {totalDuplicateMaterials}\n" +
                    $"Групп дубликатов: {totalDuplicateGroups}\n" +
                    $"Деталей с дубликатами: {detailsWithMaterialDuplicates.Count}\n" +
                    $"Отчет: {reportPath}";

                ShowMessage(call, finalMessage);
            }
            catch (Exception ex)
            {
                try { call.RunMethod("SetFormat", new object[] { "" }); } catch { }
                ShowMessage(call, $"Ошибка: {ex.Message}\n{ex.StackTrace}");
            }
        }

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
                    object res = call.RunMethod("GetAllLinkedObjects", new object[] { id.ToString(), 0 });
                    if (res is string xml && !string.IsNullOrEmpty(xml))
                    {
                        System.Diagnostics.Debug.WriteLine($"XML for all linked in {id}: {xml}");
                        var children = ParseLinksXml(xml);
                        foreach (var c in children)
                            if (!visited.Contains(c.Id))
                                queue.Enqueue(c.Id);
                    }
                }
                catch { }
            }
            return objects;
        }

        private List<ObjectInfo> GetMaterialLinks(INetPluginCall call, int objectId, int materialTypeId)
        {
            var list = new List<ObjectInfo>();
            try
            {
                object res = call.RunMethod("GetAllLinkedObjects", new object[] { objectId.ToString(), 0 });
                if (res is string xml && !string.IsNullOrEmpty(xml))
                {
                    System.Diagnostics.Debug.WriteLine($"XML for all linked in {objectId}: {xml}");
                    var all = ParseLinksXml(xml);
                    list = all.Where(o => int.TryParse(o.Type, out int t) && t == materialTypeId).ToList();

                    foreach (var o in list)
                    {
                        if (int.TryParse(o.Type, out int tid) && _typeDictionary.ContainsKey(tid))
                            o.Type = _typeDictionary[tid];
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetMaterialLinks error {objectId}: {ex.Message}");
            }
            return list;
        }

        private List<ObjectInfo> ParseLinksXml(string xmlData)
        {
            var list = new List<ObjectInfo>();
            try
            {
                var doc = new XmlDocument();
                doc.LoadXml(xmlData);
                var rows = doc.SelectNodes("//row");
                if (rows == null || rows.Count == 0) rows = doc.SelectNodes("//ROOT/rowset/row");
                if (rows == null) return list;

                foreach (XmlNode row in rows)
                {
                    var info = new ObjectInfo();
                    foreach (XmlAttribute a in row.Attributes)
                    {
                        string n = a.Name.ToUpper();
                        string v = a.Value;
                        if (n == "C0" || n == "_ID_VERSION" || n == "ID_VERSION" || n == "_ID_VERSION_CHILD")
                        {
                            int id;
                            if (int.TryParse(v, out id))
                                info.Id = id;
                        }
                        else if (n == "C1" || n == "_TYPE")
                            info.Type = v;
                        else if (n == "C2" || n == "_PRODUCT" || n == "NAME" || n == "_NAME")
                            info.Name = v;
                        else if (n == "C3" || n == "_VERSION")
                            info.Version = v;
                        else if (n == "C4" || n == "_STATE")
                            info.State = v;
                    }
                    if (info.Id > 0)
                    {
                        if (string.IsNullOrEmpty(info.Name))
                            info.Name = $"{info.Type} {info.Version}".Trim();
                        list.Add(info);
                    }
                }
            }
            catch { }
            return list;
        }

        private string DebugMaterialSearch(ObjectInfo parent, List<ObjectInfo> materials)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"=== Деталь: {parent.Name} (ID: {parent.Id}) ===");
            sb.AppendLine($"Связей с материалами: {materials.Count}");
            foreach (var m in materials)
                sb.AppendLine($" - ID {m.Id}: {m.Name} ({m.Version})");

            var groups = materials.GroupBy(m => m.Name.Trim() + " | " + m.Version.Trim())
                                  .Where(g => g.Count() > 1);

            sb.AppendLine("\nДубликаты:");
            foreach (var g in groups)
                sb.AppendLine($"   {g.Key} → {g.Count()} раз");

            return sb.ToString();
        }

        private string CreateMaterialDuplicatesHtmlReport(ObjectInfo folderInfo, int totalObjects, int totalDetails,
            int totalMaterials, int totalDups, List<DetailWithMaterialDuplicates> dups)
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

        private void ShowMessage(INetPluginCall call, string msg)
        {
            try { call.RunMethod("ShowMessage", new object[] { msg }); }
            catch
            {
                try { call.RunMethod("ShowErrorMessage", new object[] { msg }); }
                catch
                {
                    try { call.RunMethod("ShowInfoMessage", new object[] { msg }); }
                    catch { System.Diagnostics.Debug.WriteLine(msg); }
                }
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
                    if (rows == null) return dict;

                    foreach (XmlNode r in rows)
                    {
                        int id = 0; string name = "";
                        foreach (XmlAttribute a in r.Attributes)
                        {
                            string n = a.Name.ToUpper();
                            if (n == "C0" || n == "_ID" || n == "ID") int.TryParse(a.Value, out id);
                            if (n == "C1" || n == "_TYPENAME" || n == "TYPENAME") name = a.Value;
                        }
                        if (id > 0 && !string.IsNullOrEmpty(name) && !dict.ContainsKey(id))
                            dict[id] = name;
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
                    var ids = s.Split(',');
                    if (ids.Length > 0 && int.TryParse(ids[0].Trim(), out int id) && id > 0)
                        return id;
                }
            }
            catch { }
            return 0;
        }

        private ObjectInfo GetObjectInfo(INetPluginCall call, int id)
        {
            try
            {
                object res = call.RunMethod("GetPropObjects", new object[] { id.ToString(), 0 });
                if (res is string xml && !string.IsNullOrEmpty(xml))
                    return ParseObjectInfoXml(xml, id);
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
                        string v = a.Value;
                        if (n == "C1" || n == "_TYPE") info.Type = v;
                        if (n == "C2" || n == "_PRODUCT") info.Name = v;
                        if (n == "C3" || n == "_VERSION") info.Version = v;
                        if (n == "C4" || n == "_STATE") info.State = v;
                    }
                    if (string.IsNullOrEmpty(info.Name))
                        info.Name = $"{info.Type} {info.Version}".Trim();
                    return info;
                }
            }
            catch { }
            return new ObjectInfo { Id = id, Name = $"Объект_{id}" };
        }
    }

    public class DetailWithMaterialDuplicates
    {
        public ObjectInfo Detail { get; set; }
        public List<MaterialGroup> MaterialGroups { get; set; }
    }

    public class MaterialGroup
    {
        public string Key { get; set; }
        public string MaterialName { get; set; } = "";
        public string MaterialType { get; set; } = "";
        public int LinkCount { get; set; }
        public List<ObjectInfo> Materials { get; set; } = new List<ObjectInfo>();
    }

    public class ObjectInfo
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string Version { get; set; } = "";
        public string State { get; set; } = "";
    }
}