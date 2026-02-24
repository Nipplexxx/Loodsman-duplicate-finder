using Ascon.Plm.Loodsman.PluginSDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;

namespace DeepDuplicateFinder
{
    [LoodsmanPlugin]
    public class DeepDuplicateFinder : ILoodsmanNetPlugin
    {
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr h, string m, string c, int type);

        private string LogPath => Path.Combine(Path.GetTempPath(), "DeepDuplicateFinder.log");
        private void Log(string message) => File.AppendAllText(LogPath, $"{DateTime.Now:HH:mm:ss.fff}: {message}\n");

        private LoodsmanApiHelper _api;
        private Dictionary<int, string> _typeDictionary;
        private int _preparationTypeId;
        private int _preparationLinkTypeId;

        public void PluginLoad() => Log("Плагин загружен");
        public void PluginUnload() { }
        public void OnConnectToDb(INetPluginCall call) { }
        public void OnCloseDb() { }

        public void BindMenu(IMenuDefinition menu)
        {
            menu.AddMenuItem("Найти множественные материалы (детали+заготовки)", Main, call => true);
        }

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

        private void Main(INetPluginCall call)
        {
            try
            {
                // Устанавливаем формат XML
                try 
                { 
                    call.RunMethod("SetFormat", new object[] { "xml" }); 
                } 
                catch 
                { 
                    
                }

                int selectedId = GetSelectedId(call);
                if (selectedId == 0)
                {
                    return;
                }

                _api = new LoodsmanApiHelper(call, Log);
                
                // Словарь типов
                _typeDictionary = _api.GetTypeDictionary();

                int detailTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(PluginConstants.DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                _preparationTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(PluginConstants.PREPARATION_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                int materialTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(PluginConstants.MATERIAL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;
                int mainMaterialTypeId = _typeDictionary.FirstOrDefault(kv => kv.Value.Equals(PluginConstants.MATERIAL_MAIN_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).Key;

                Log($"ID детали: {detailTypeId}, ID заготовки: {_preparationTypeId}");

                // Собираем все объекты из папки
                List<ObjectInfo> allObjects = _api.GetAllObjectsIncludingIndirect(selectedId);

                // Определяем ID связи "Заготовка для"
                _preparationLinkTypeId = _api.FindPreparationLinkTypeId(allObjects, PluginConstants.PREPARATION_LINK_NAME);
                Log($"ID связи '{PluginConstants.PREPARATION_LINK_NAME}': {_preparationLinkTypeId}");

                // Добавляем недостающие заготовки, найденные через связи
                if (_preparationLinkTypeId != 0)
                {
                    var extraPreps = _api.FindAllPreparations(allObjects, _preparationLinkTypeId, PluginConstants.PREPARATION_TYPE_NAME);
                    foreach (var prep in extraPreps)
                    {
                        if (!allObjects.Any(o => o.Id == prep.Id))
                        {
                            allObjects.Add(prep);
                            Log($"Добавлена заготовка через связь: ID={prep.Id}, Name={prep.Name}");
                        }
                    }
                }

                // Отделяем детали и заготовки
                var details = allObjects.Where(o => o.Type.Equals(PluginConstants.DETAIL_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
                var preparations = allObjects.Where(o => o.Type.Equals(PluginConstants.PREPARATION_TYPE_NAME, StringComparison.OrdinalIgnoreCase)).ToList();
                Log($"Всего объектов: {allObjects.Count}, деталей: {details.Count}, заготовок: {preparations.Count}");

                var reportItems = new List<ReportItem>();
                var allDebugInfo = new StringBuilder();
                int totalMaterialsFound = 0;

                // Анализ деталей (материалы по КД)
                if (materialTypeId != 0)
                {
                    foreach (var detail in details)
                    {
                        var materialLinks = _api.GetMaterialLinks(detail.Id, materialTypeId, _typeDictionary);
                        if (materialLinks.Count > 0)
                        {
                            totalMaterialsFound += materialLinks.Count;
                            allDebugInfo.AppendLine(DebugMaterialSearch(detail, materialLinks, "Деталь", PluginConstants.MATERIAL_TYPE_NAME));
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
                                    MaterialTypeName = PluginConstants.MATERIAL_TYPE_NAME
                                });
                            }
                        }
                    }
                }

                // Анализ заготовок (основные материалы)
                if (_preparationTypeId != 0 && mainMaterialTypeId != 0)
                {
                    foreach (var preparation in preparations)
                    {
                        var materialLinks = _api.GetMaterialLinks(preparation.Id, mainMaterialTypeId, _typeDictionary);
                        if (materialLinks.Count > 0)
                        {
                            totalMaterialsFound += materialLinks.Count;
                            allDebugInfo.AppendLine(DebugMaterialSearch(preparation, materialLinks, "Заготовка", PluginConstants.MATERIAL_MAIN_TYPE_NAME));
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
                                    MaterialTypeName = PluginConstants.MATERIAL_MAIN_TYPE_NAME
                                });
                            }
                        }
                    }
                }

                // Собираем детали с несколькими заготовками
                var detailsWithMultiplePreparations = new List<ObjectInfo>();
                if (_preparationLinkTypeId != 0)
                {
                    foreach (var detail in details)
                    {
                        int prepLinksCount = _api.GetPreparationLinksCount(detail.Id, _preparationLinkTypeId);
                        if (prepLinksCount > 1)
                        {
                            detailsWithMultiplePreparations.Add(detail);
                            Log($"Деталь с несколькими заготовками: ID={detail.Id}, Name={detail.Name}, связей={prepLinksCount}");
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

                Log("--------------------------Отладка--------------------------");
                Log(allDebugInfo.ToString());

                // Формируем список ID для открытия – все проблемные объекты (и детали, и заготовки)
                var idsToOpen = new HashSet<int>();
                foreach (var item in reportItems)
                {
                    idsToOpen.Add(item.Object.Id);
                }
                foreach (var detail in detailsWithMultiplePreparations)
                {
                    idsToOpen.Add(detail.Id);
                }

                Log($"Всего объектов для открытия (только детали): {idsToOpen.Count}");

                // Открываем окна, если есть что открывать
                if (idsToOpen.Count > 0)
                {
                    string idsString = string.Join(",", idsToOpen);
                    var opener = new WindowOpener();
                    opener.ShowObjectsInNewWindows(call, idsString);
                }

                // Показываем статистику всегда, если есть проблемы
                bool hasProblems = reportItems.Count > 0 || detailsWithMultiplePreparations.Count > 0;
                if (hasProblems)
                {
                    string stats = "";
                    if (reportItems.Any(r => r.ObjectTypeName == "Деталь"))
                        stats += $"\nДеталей с >1 материалом: {reportItems.Count(r => r.ObjectTypeName == "Деталь")}";
                    if (reportItems.Any(r => r.ObjectTypeName == "Заготовка"))
                        stats += $"\nЗаготовок с >1 материалом: {reportItems.Count(r => r.ObjectTypeName == "Заготовка")}";
                    if (detailsWithMultiplePreparations.Count > 0)
                        stats += $"\nДеталей с несколькими заготовками: {detailsWithMultiplePreparations.Count}";

                    string finalMessage = $"Найдено проблемных объектов:{stats}";
                    MessageBox((IntPtr)0, finalMessage, "DeepDuplicateFinder", 0);
                }
                else
                {
                    MessageBox((IntPtr)0, "Объекты с множественными материалами или заготовками не найдены.", "DeepDuplicateFinder", 0);
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка: {ex}");
                MessageBox((IntPtr)0, $"Ошибка: {ex.Message}", "DeepDuplicateFinder", 0);
                try 
                { 
                    call.RunMethod("SetFormat", new object[] { "" }); 
                } 
                catch 
                { 
                    
                }
            }
        }

        private string DebugMaterialSearch(ObjectInfo obj, List<ObjectInfo> materials, string objTypeName, string materialTypeName)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{objTypeName}: {obj.Name} (ID: {obj.Id}, Версия: {obj.Version})");
            sb.AppendLine($"Тип материала: {materialTypeName}");
            sb.AppendLine($"Связей: {materials.Count}");
            foreach (var m in materials)
                sb.AppendLine($" - ID {m.Id}: {m.Name} (Версия: {m.Version})");
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