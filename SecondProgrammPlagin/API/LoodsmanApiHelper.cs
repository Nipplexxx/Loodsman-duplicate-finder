using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Ascon.Plm.Loodsman.PluginSDK;

namespace DeepDuplicateFinder
{
    public class LoodsmanApiHelper
    {
        private readonly INetPluginCall _call;
        private readonly Action<string> _log;

        public LoodsmanApiHelper(INetPluginCall call, Action<string> log = null)
        {
            _call = call;
            _log = log ?? (msg => { });
        }

        public Dictionary<int, string> GetTypeDictionary()
        {
            var dict = new Dictionary<int, string>();
            try
            {
                object res = _call.RunMethod("GetTypeList", new object[0]);
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
                _log($"GetTypeDictionary ошибка: {ex.Message}"); 
            }
            return dict;
        }

        public ObjectInfo GetObjectInfo(int id)
        {
            var result = new ObjectInfo { Id = id, Name = $"Объект_{id}", Type = "Неизвестный", Version = "", State = "" };
            try
            {
                object res = _call.RunMethod("GetPropObjects", new object[] { id.ToString(), 0 });
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

        public List<LinkedObjectInfo> GetLinkedObjects(int objectId, bool reverse)
        {
            var list = new List<LinkedObjectInfo>();
            try
            {
                object res = _call.RunMethod("GetLinkedObjectsForObjects", new object[] { objectId.ToString(), "", reverse });
                if (res is string xml && !string.IsNullOrEmpty(xml))
                {
                    list = ParseLinkedObjectsXml(xml);
                }
            }
            catch (Exception ex) 
            {
                _log($"GetLinkedObjects ошибка для {objectId}: {ex.Message}"); 
            }
            return list;
        }

        public List<ObjectInfo> GetMaterialLinks(int objectId, int materialTypeId, Dictionary<int, string> typeDict)
        {
            var list = new List<ObjectInfo>();
            try
            {
                var linked = GetLinkedObjects(objectId, false);
                var materialLinks = linked.Where(o => o.TypeId == materialTypeId).ToList();

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
                    if (int.TryParse(o.Type, out int tid) && typeDict.ContainsKey(tid))
                    {
                        o.Type = typeDict[tid];
                    }
                }
            }
            catch (Exception ex) 
            { 
                _log($"GetMaterialLinks ошибка для {objectId}: {ex.Message}"); 
            }
            return list;
        }

        public int GetPreparationLinksCount(int detailId, int preparationLinkTypeId)
        {
            try
            {
                var linked = GetLinkedObjects(detailId, true);
                return linked.Count(l => l.LinkTypeId == preparationLinkTypeId);
            }
            catch (Exception ex)
            {
                _log($"GetPreparationLinksCount ошибка для детали {detailId}: {ex.Message}");
                return 0;
            }
        }

        public List<ObjectInfo> FindAllPreparations(List<ObjectInfo> allObjects, int preparationLinkTypeId, string preparationTypeName)
        {
            var result = new List<ObjectInfo>();
            try
            {
                var details = allObjects.Where(o => o.Type.Equals("Деталь", StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (var detail in details)
                {
                    var linked = GetLinkedObjects(detail.Id, true);
                    var prepLinks = linked.Where(l => l.LinkTypeId == preparationLinkTypeId).ToList();
                    foreach (var link in prepLinks)
                    {
                        var prepInfo = GetObjectInfo(link.ChildId);
                        if (prepInfo.Id > 0 && prepInfo.Type.Equals(preparationTypeName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (!result.Any(r => r.Id == prepInfo.Id))
                            {
                                result.Add(prepInfo);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) 
            { 
                _log($"FindAllPreparations ошибка: {ex.Message}"); 
            }
            return result;
        }

        public int FindPreparationLinkTypeId(List<ObjectInfo> allObjects, string preparationLinkName)
        {
            var details = allObjects.Where(o => o.Type.Equals("Деталь", StringComparison.OrdinalIgnoreCase)).ToList();
            _log($"Поиск ID связи '{preparationLinkName}': найдено деталей {details.Count}");
            foreach (var detail in details)
            {
                try
                {
                    var linked = GetLinkedObjects(detail.Id, true);
                    var prepLink = linked.FirstOrDefault(l => l.LinkTypeName?.Trim().Equals(preparationLinkName, StringComparison.OrdinalIgnoreCase) == true);
                    if (prepLink != null)
                    {
                        _log($"Найден ID связи '{preparationLinkName}': {prepLink.LinkTypeId} (по объекту {detail.Id})");
                        return prepLink.LinkTypeId;
                    }
                }
                catch (Exception ex) 
                { 
                    _log($"Ошибка при поиске связи для объекта {detail.Id}: {ex.Message}"); 
                }
            }
            _log($"Не удалось найти ID связи '{preparationLinkName}'");
            return 0;
        }

        public List<ObjectInfo> GetAllObjectsIncludingIndirect(int rootId)
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

                var info = GetObjectInfo(id);
                if (info.Id > 0 && !objects.Any(o => o.Id == info.Id))
                    objects.Add(info);

                try
                {
                    var children = GetLinkedObjects(id, false);
                    foreach (var c in children)
                        if (!visited.Contains(c.ChildId))
                            queue.Enqueue(c.ChildId);

                    var parents = GetLinkedObjects(id, true);
                    foreach (var p in parents)
                        if (!visited.Contains(p.ParentId) && p.ParentId != id)
                            queue.Enqueue(p.ParentId);
                }
                catch (Exception ex) { _log($"Ошибка при обходе объекта {id}: {ex.Message}"); }
            }
            return objects;
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
                        string val = a.Value.Trim();
                        if (n == "C1" || n == "_TYPE") info.Type = val;
                        if (n == "C2" || n == "_PRODUCT") info.Name = val;
                        if (n == "C3" || n == "_VERSION") info.Version = val;
                        if (n == "C4" || n == "_STATE") info.State = val;
                    }
                    if (string.IsNullOrEmpty(info.Name))
                        info.Name = $"{info.Type} {info.Id}".Trim();
                    return info;
                }
            }
            catch { }
            return new ObjectInfo { Id = id, Name = $"Объект_{id}" };
        }

        private List<LinkedObjectInfo> ParseLinkedObjectsXml(string xmlData)
        {
            var list = new List<LinkedObjectInfo>();
            try
            {
                var doc = new XmlDocument();
                doc.LoadXml(xmlData);
                var rows = doc.SelectNodes("//row") ?? doc.SelectNodes("//ROOT/rowset/row");
                if (rows == null) return list;

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
                    if (info.ChildId > 0) list.Add(info);
                }
            }
            catch (Exception ex) { _log($"ParseLinkedObjectsXml ошибка: {ex.Message}"); }
            return list;
        }
    }
}