using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassicVolumeMixer.Helpers
{
    public static class ThemeHelper
    {
        public static bool SystemUsesLightTheme()
        {
            const string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                if (key != null)
                {
                    object value = key.GetValue("SystemUsesLightTheme");
                    if (value is int systemUsesLightTheme)
                    {
                        return systemUsesLightTheme == 1;
                    }
                }
            }
            return false;
        }
    }
}
