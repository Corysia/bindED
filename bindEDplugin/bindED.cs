using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace bindEDplugin
{
    public class bindEDPlugin
    {
        private static readonly Dictionary<string, int> _map = new Dictionary<string, int>(256);

        public static string VA_DisplayName()
        {
            return "bindED Plugin v1.1";
        }

        public static string VA_DisplayInfo()
        {
            return "bindED Plugin\r\n\r\n2016 VoiceAttack.com";
        }

        public static Guid VA_Id()
        {
            return new Guid("{53D74F33-3F5C-487B-B6D2-A17593DEBE05}");
        }

        public static void VA_Init1(dynamic vaProxy)
        {
            LoadBinds(vaProxy, false);
        }

        private static string GetPluginPath(dynamic vaProxy)
        {
            return Path.GetDirectoryName(vaProxy.PluginPath());
        }

        private static void LoadBinds(dynamic vaProxy, bool fromInvoke)
        {
            string strDir = GetPluginPath(vaProxy);
            var strMap = Path.Combine(strDir, "EDMap.txt");
            try
            {
                if (File.Exists(strMap))
                {
                    foreach (var line in File.ReadAllLines(strMap, Encoding.UTF8))
                    {
                        var arItem = line.Split(";".ToCharArray(), 2, StringSplitOptions.RemoveEmptyEntries);
                        if (arItem.Count() != 2 || string.IsNullOrWhiteSpace(arItem[0]) ||
                            _map.ContainsKey(arItem[0])) continue;
                        ushort iKey;
                        if (!ushort.TryParse(arItem[1], out iKey)) continue;
                        if (iKey > 0 && iKey < 256)
                            _map.Add(arItem[0].Trim(), iKey);
                    }
                }
                else
                {
                    vaProxy.WriteToLog(
                        "bindED Error - EDMap.txt does not exist.  Make sure the EDMap.txt file exists in the same folder as this plugin, otherwise this plugin has nothing to process and cannot continue.",
                        "red");
                    return;
                }
            }
            catch (Exception ex)
            {
                vaProxy.WriteToLog($"bindED Error - {ex.Message}", "red");
                return;
            }

            if (_map.Count == 0)
            {
                vaProxy.WriteToLog("bindED Error - EDMap.txt does not contain any elements.", "red");
                return;
            }

            string[] files = null;

            if (fromInvoke)
            {
                var strBindsDir =
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        @"Frontier Developments\Elite Dangerous\Options\Bindings");
                if (!string.IsNullOrWhiteSpace(vaProxy.Context))
                {
                    files = ((string) vaProxy.Context).Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                }
                else
                {
                    if (Directory.Exists(strBindsDir))
                    {
                        var bindFiles = new DirectoryInfo(strBindsDir).GetFiles().Where(i => i.Extension == ".binds")
                            .OrderByDescending(p => p.LastWriteTime).ToArray();

                        if (bindFiles.Any())
                            files = new[] {bindFiles[0].FullName};
                    }
                }
            }
            else
            {
                try
                {
                    files = Directory.GetFiles(strDir, "*.lnk", SearchOption.TopDirectoryOnly);
                    IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                    for (var i = 0; i < files.Count(); i++)
                        files[i] = ((IWshRuntimeLibrary.IWshShortcut) shell.CreateShortcut(files[i])).TargetPath;
                }
                catch (Exception ex)
                {
                    vaProxy.WriteToLog($"bindED Error - {ex.Message}", "red");
                    return;
                }
            }

            try
            {
                foreach (var file in files)
                    if (File.Exists(file))
                    {
                        XElement rootElement = null;

                        try
                        {
                            rootElement = XElement.Load(file);
                        }
                        catch (Exception ex)
                        {
                            vaProxy.WriteToLog($"bindED Error - {ex.Message}", "red");
                            return;
                        }

                        if (rootElement == null) continue;
                        foreach (var c in rootElement.Elements().Where(i => i.Elements().Any()))
                        foreach (var element in c.Elements().Where(i => i.HasAttributes))
                        {
                            var _keys = new List<int>();
                            if (element.Name == "Primary")
                                if (element.Attribute("Device").Value == "Keyboard" &&
                                    !string.IsNullOrWhiteSpace(element.Attribute("Key").Value) &&
                                    element.Attribute("Key").Value.StartsWith("Key_"))
                                {
                                    _keys.AddRange(
                                        from modifier in element.Elements()
                                            .Where(i => i.Name.LocalName == "Modifier")
                                        where _map.ContainsKey(modifier.Attribute("Key").Value)
                                        select _map[modifier.Attribute("Key").Value]);

                                    if (_map.ContainsKey(element.Attribute("Key").Value))
                                        _keys.Add(_map[element.Attribute("Key").Value]);
                                }

                            if (_keys.Count == 0) //nothing found in primary... look in secondary
                                if (element.Name == "Secondary")
                                    if (element.Attribute("Device").Value == "Keyboard" &&
                                        !string.IsNullOrWhiteSpace(element.Attribute("Key").Value) &&
                                        element.Attribute("Key").Value.StartsWith("Key_"))
                                    {
                                        _keys.AddRange(
                                            from modifier in element.Elements()
                                                .Where(i => i.Name.LocalName == "Modifier")
                                            where _map.ContainsKey(modifier.Attribute("Key").Value)
                                            select _map[modifier.Attribute("Key").Value]);

                                        if (_map.ContainsKey(element.Attribute("Key").Value))
                                            _keys.Add(_map[element.Attribute("Key").Value]);
                                    }

                            if (_keys.Count <= 0) continue;
                            var strTextValue = string.Empty;
                            foreach (var key in _keys)
                                strTextValue += $"[{key}]";

                            vaProxy.SetText($"ed{c.Name.LocalName}", strTextValue);
                        }
                    }
                    else
                    {
                        vaProxy.WriteToLog(
                            $"bindED Error - The target file indicated by the shortcut does not exist: {file}", "red");
                        return;
                    }
            }
            catch (Exception ex)
            {
                vaProxy.WriteToLog($"bindED Error - {ex.Message}", "red");
            }
        }

        public static void VA_Exit1(dynamic vaProxy)
        {
        }

        public static void VA_StopCommand()
        {
        }

        public static void VA_Invoke1(dynamic vaProxy)
        {
            LoadBinds(vaProxy, true);
        }
    }
}