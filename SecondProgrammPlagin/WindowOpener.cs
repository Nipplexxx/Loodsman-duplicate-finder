using System;
using System.Runtime.InteropServices;
using Ascon.Plm.Loodsman.PluginSDK;

namespace DeepDuplicateFinder
{
    public class WindowOpener
    {
        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(IntPtr h, string m, string c, int type);

        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(HandleRef hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Ansi)]
        private struct TMSGPARAMS
        {
            public IntPtr Reserved;      // должен быть null
            public string CheckOutName;   // имя чекаута или пустая строка (ANSI)
            public uint ObjectId;          // идентификатор одиночного объекта
            public string ObjectIds;       // список ID через запятую (ANSI)
        }

        public void ShowObjectsInNewWindows(INetPluginCall pluginCall, string objectIds)
        {
            if (string.IsNullOrEmpty(objectIds))
            {
                MessageBox(IntPtr.Zero, "Нет объектов для отображения.", "DeepDuplicateFinder", 0);
                return;
            }

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

            TMSGPARAMS param = new TMSGPARAMS
            {
                Reserved = IntPtr.Zero,
                CheckOutName = null, // null означает отсутствие чекаута
                ObjectId = 0,
                ObjectIds = objectIds
            };

            int size = Marshal.SizeOf(param);
            IntPtr ptr = Marshal.AllocHGlobal(size);
            try
            {
                Marshal.StructureToPtr(param, ptr, false);

                const uint WM_OPENOBJECTSINNEWWINDOW = 0x0400 + 101;
                IntPtr result = SendMessage(new HandleRef(null, mainHandle),
                                            WM_OPENOBJECTSINNEWWINDOW,
                                            ptr,
                                            IntPtr.Zero);
                // Если результат не 0, возможно, ошибка
                if (result != IntPtr.Zero)
                {
                    // В документации не сказано, что означает возвращаемое значение.
                    // Можно залогировать, но не прерываем.
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DeepDuplicateFinder.log");
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: SendMessage returned {result}\n");
                }
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }
    }
}