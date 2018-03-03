﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;

namespace bindEDplugin
{
    public class bindEDPlugin
    {
        private static Dictionary<String, int> _map = new Dictionary<string, int>(256);
        public static string VA_DisplayName()
        {
            return "bindED Plugin v1.0";  
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

        private static String GetPluginPath(dynamic vaProxy)
        {
            return Path.GetDirectoryName(vaProxy.PluginPath());
        }

        public static void LoadBinds(dynamic vaProxy, Boolean fromInvoke)
        {
            String strDir = GetPluginPath(vaProxy);
            try
            {
                String strMap = Path.Combine(strDir, "EDMap.txt");
                if (File.Exists(strMap))
                {
                    foreach (String line in File.ReadAllLines(strMap, System.Text.Encoding.UTF8))
                    {
                        String[] arItem = line.Split(";".ToCharArray(), 2, StringSplitOptions.RemoveEmptyEntries);
                        if ((arItem.Count() == 2) && (!String.IsNullOrWhiteSpace(arItem[0])) && (!_map.ContainsKey(arItem[0])))
                        {
                            ushort iKey;
                            if (ushort.TryParse(arItem[1], out iKey))
                            {
                                if (iKey > 0 && iKey < 256)
                                    _map.Add(arItem[0].Trim(), iKey);
                            }
                        }
                    }
                }
                else
                {
                    vaProxy.WriteToLog("bindED Error - EDMap.txt does not exist.  Make sure the EDMap.txt file exists in the same folder as this plugin, otherwise this plugin has nothing to process and cannot continue.", "red");
                    return;
                }
            }
            catch (Exception ex)
            {
                vaProxy.WriteToLog("bindED Error - " + ex.Message, "red");
                return;
            }

            if (_map.Count == 0)
            {
                vaProxy.WriteToLog("bindED Error - EDMap.txt does not contain any elements.", "red");
                return;
            }

            String[] files = null;

            if (fromInvoke)
            {
                if (!String.IsNullOrWhiteSpace(vaProxy.Context))
                {
                    files = ((String)vaProxy.Context).Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                }
                else
                {
                    String strBindsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Frontier Developments\Elite Dangerous\Options\Bindings");
                    if (System.IO.Directory.Exists(strBindsDir))
                    {
                        FileInfo[] bindFiles = new DirectoryInfo(strBindsDir).GetFiles().Where(i => i.Extension == ".binds").OrderByDescending(p => p.LastWriteTime).ToArray();

                        if (bindFiles.Count() > 0)
                            files = new string[] { bindFiles[0].FullName };
                    }
                }
            }
            else
            {
                try
                {
                    files = System.IO.Directory.GetFiles(strDir, "*.lnk", SearchOption.TopDirectoryOnly);
                    IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                    for (int i = 0; i < files.Count(); i++)
                    {
                        files[i] = ((IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(files[i])).TargetPath;
                    }
                }
                catch (Exception ex)
                {
                    vaProxy.WriteToLog("bindED Error - " + ex.Message, "red");
                    return;
                }
            }
            try
            {
                foreach (String file in files)
                {
                    if (File.Exists(file))
                    {

                        XElement rootElement = null;

                        try
                        {
                            rootElement = XElement.Load(file);
                        }
                        catch (Exception ex)
                        {
                            vaProxy.WriteToLog("bindED Error - " + ex.Message, "red");
                            return;
                        }

                        if (rootElement != null)
                        {

                            foreach (XElement c in rootElement.Elements().Where(i => i.Elements().Count() > 0))
                            {
                                foreach (var element in c.Elements().Where(i => i.HasAttributes))
                                {
                                    List<int> _keys = new List<int>();
                                    if (element.Name == "Primary")
                                    {
                                        if (element.Attribute("Device").Value == "Keyboard" && !String.IsNullOrWhiteSpace(element.Attribute("Key").Value) && element.Attribute("Key").Value.StartsWith("Key_"))
                                        {
                                            foreach (var modifier in element.Elements().Where(i => i.Name.LocalName == "Modifier"))
                                            {
                                                if (_map.ContainsKey(modifier.Attribute("Key").Value))
                                                    _keys.Add(_map[modifier.Attribute("Key").Value]);
                                            }

                                            if (_map.ContainsKey(element.Attribute("Key").Value))
                                                _keys.Add(_map[element.Attribute("Key").Value]);
                                        }
                                    }
                                    if (_keys.Count == 0) //nothing found in primary... look in secondary
                                    {
                                        if (element.Name == "Secondary")
                                        {
                                            if (element.Attribute("Device").Value == "Keyboard" && !String.IsNullOrWhiteSpace(element.Attribute("Key").Value) && element.Attribute("Key").Value.StartsWith("Key_"))
                                            {
                                                foreach (var modifier in element.Elements().Where(i => i.Name.LocalName == "Modifier"))
                                                {
                                                    if (_map.ContainsKey(modifier.Attribute("Key").Value))
                                                        _keys.Add(_map[modifier.Attribute("Key").Value]);
                                                }

                                                if (_map.ContainsKey(element.Attribute("Key").Value))
                                                    _keys.Add(_map[element.Attribute("Key").Value]);
                                            }
                                        }
                                    }

                                    if (_keys.Count > 0)
                                    {
                                        String strTextValue = String.Empty;
                                        foreach (int key in _keys)
                                            strTextValue += String.Format("[{0}]", key);

                                        vaProxy.SetText("ed" + c.Name.LocalName, strTextValue);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        vaProxy.WriteToLog("bindED Error - The target file indicated by the shortcut does not exist: " + file, "red");
                        return;
                    }
                }
            }
            catch(Exception ex)
            {
                vaProxy.WriteToLog("bindED Error - " + ex.Message, "red");
                return;
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
