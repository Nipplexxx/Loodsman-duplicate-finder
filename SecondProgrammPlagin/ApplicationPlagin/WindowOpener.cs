using System;
using System.Runtime.InteropServices;
using Ascon.Plm.Loodsman.PluginSDK;

namespace DeepDuplicateFinder
{
    // Класс для открытия новых окон ЛОЦМАН с переданными объектами.
    public class WindowOpener
    {
        // Импорт функции MessageBox для отображения сообщений об ошибках
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(IntPtr h, string m, string c, int type);

        // Импорт функции SendMessage для отправки сообщения главному окну ЛОЦМАН
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(HandleRef hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Ansi)]
        private struct TMSGPARAMS
        {
            public IntPtr Reserved;
            public string CheckOutName;   // имя чекаута
            public uint ObjectId;         // идентификатор одиночного объекта
            public string ObjectIds;      // строка с идентификаторами объектов через запятую
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

            // Получаем MainHandle через PluginCall
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
                CheckOutName = null,
                ObjectId = 0,
                ObjectIds = objectIds      // передаём список ID
            };

            // Выделяем память под структуру
            int size = Marshal.SizeOf(param);
            IntPtr ptr = Marshal.AllocHGlobal(size);
            try
            {
                // Копируем структуру
                Marshal.StructureToPtr(param, ptr, false);

                const uint WM_OPENOBJECTSINNEWWINDOW = 0x0400 + 101;

                // Отправляем сообщение главному окну ЛОЦМАН
                IntPtr result = SendMessage(new HandleRef(null, mainHandle),
                                            WM_OPENOBJECTSINNEWWINDOW,
                                            ptr,
                                            IntPtr.Zero);

                if (result != IntPtr.Zero)
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DeepDuplicateFinder.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: SendMessage returned {result}\n");
                }
            }
            finally
            {
                // Освобождаем выделенную память
                Marshal.FreeHGlobal(ptr);
            }
        }
    }
}