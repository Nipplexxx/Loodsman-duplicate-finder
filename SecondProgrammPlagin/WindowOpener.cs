using System;
using System.Runtime.InteropServices;
using Ascon.Plm.Loodsman.PluginSDK;

namespace DeepDuplicateFinder
{
    // Класс для открытия новых окон ЛОЦМАН с переданными объектами.
    // Использует оконное сообщение WM_OPENOBJECTSINNEWWINDOW (WM_USER + 101).
    public class WindowOpener
    {
        // Импорт функции MessageBox для отображения сообщений об ошибках
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(IntPtr h, string m, string c, int type);

        // Импорт функции SendMessage для отправки сообщения главному окну ЛОЦМАН
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(HandleRef hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        // Структура, соответствующая параметрам сообщения WM_OPENOBJECTSINNEWWINDOW.
        [StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Ansi)]
        private struct TMSGPARAMS
        {
            public IntPtr Reserved;      // зарезервировано, должен быть null (IntPtr.Zero)
            public string CheckOutName;   // имя чекаута, если окно нужно открыть в контексте чекаута; иначе null
            public uint ObjectId;         // идентификатор одиночного объекта (используется, если ObjectIds == null)
            public string ObjectIds;      // строка с идентификаторами объектов через запятую (например, "616,2790")
        }

        // Открывает новые окна ЛОЦМАН для списка объектов, идентификаторы которых переданы в objectIds.
        public void ShowObjectsInNewWindows(INetPluginCall pluginCall, string objectIds)
        {
            // Проверка входных данных
            if (string.IsNullOrEmpty(objectIds))
            {
                MessageBox(IntPtr.Zero, "Нет объектов для отображения.", "DeepDuplicateFinder", 0);
                return;
            }

            // Получаем MainHandle через PluginCall (свойство PluginCall интерфейса INetPluginCall)
            if (pluginCall?.PluginCall == null)
            {
                MessageBox(IntPtr.Zero, "Не удалось получить MainHandle.", "DeepDuplicateFinder", 0);
                return;
            }

            IntPtr mainHandle = (IntPtr)pluginCall.PluginCall.MainHandle;
            if (mainHandle == IntPtr.Zero)
            {
                MessageBox(IntPtr.Zero, "MainHandle = 0.", "DeepDuplicateFinder", 0);
                return;
            }

            // Заполняем структуру параметров
            TMSGPARAMS param = new TMSGPARAMS
            {
                Reserved = IntPtr.Zero,
                CheckOutName = null,      // не используем чекаут
                ObjectId = 0,              // не используем одиночный объект
                ObjectIds = objectIds      // передаём список ID
            };

            // Выделяем память под структуру в неуправляемой куче
            int size = Marshal.SizeOf(param);
            IntPtr ptr = Marshal.AllocHGlobal(size);
            try
            {
                // Копируем структуру в выделенную память
                Marshal.StructureToPtr(param, ptr, false);

                // Идентификатор сообщения: WM_USER + 101
                const uint WM_OPENOBJECTSINNEWWINDOW = 0x0400 + 101;

                // Отправляем сообщение главному окну ЛОЦМАН
                IntPtr result = SendMessage(new HandleRef(null, mainHandle),
                                            WM_OPENOBJECTSINNEWWINDOW,
                                            ptr,          // wParam – указатель на структуру TMSGPARAMS
                                            IntPtr.Zero); // lParam – не используется

                // Если результат не 0, возможно, ошибка – логируем (но не прерываем выполнение)
                if (result != IntPtr.Zero)
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DeepDuplicateFinder.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: SendMessage returned {result}\n");
                }
            }
            finally
            {
                // Обязательно освобождаем выделенную память
                Marshal.FreeHGlobal(ptr);
            }
        }
    }
}