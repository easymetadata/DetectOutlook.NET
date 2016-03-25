using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Microsoft.Win32;

namespace DetectOutlook.NET
{
    public class outlookdetect
    {

        public static string STRbitness = string.Empty;

        public static void isOutlook64bit()
        {
            try
            {
                List<String> _SourceKeys = new List<string>();
                _SourceKeys.Add(@"SOFTWARE\Microsoft\Office");
                _SourceKeys.Add(@"SOFTWARE\Wow6432Node\Microsoft\Office");

                foreach (string _sourceKey in _SourceKeys)
                {
                    RegistryKey items = Registry.LocalMachine.OpenSubKey(_sourceKey);

                    foreach (string _s in items.GetSubKeyNames().ToList())
                    {
                        if (_s.Contains(".") && String.IsNullOrEmpty(STRbitness))
                        {
                            string _s2 = Path.Combine(_sourceKey, _s, "Outlook");
                            string _s3 = Path.Combine(_s2, "InstallRoot");

                            try
                            {
                                RegistryKey subitem = Registry.LocalMachine.OpenSubKey(@_s2);

                                if (subitem != null)
                                {
                                    STRbitness = (string)subitem.GetValue("Bitness", String.Empty);

                                    Console.WriteLine("\nOutlook info");
                                    Console.WriteLine("-----------------------");
                                    Console.WriteLine("Reg Key:     " + Path.Combine(_sourceKey, _s2));
                                    Console.WriteLine("Bitness:     " + STRbitness);

                                }

                                if (!string.IsNullOrEmpty(STRbitness))
                                {
                                    RegistryKey subitem2 = Registry.LocalMachine.OpenSubKey(@_s3);

                                    if (subitem2 != null)
                                    {
                                        var path = (string)subitem2.GetValue("Path", String.Empty);
                                        Console.WriteLine("Path:        " + path);
                                    }
                                }

                                if (!string.IsNullOrEmpty(STRbitness))
                                {
                                    string _s4 = Path.Combine(_sourceKey, _s, "Common", "ProductVersion");
                                    RegistryKey subitem3 = Registry.LocalMachine.OpenSubKey(@_s4);

                                    if (subitem3 != null)
                                    {
                                        var path = (string)subitem3.GetValue("LastProduct", String.Empty);
                                        Console.WriteLine("Version:     " + path);
                                    }
                                }

                            }
                            catch (Exception) { }
                        }

                    }
                }

                if (string.IsNullOrEmpty(STRbitness))
                    Console.WriteLine("Outlook not detected.");

            }
            catch (Exception)
            {
            }

        }

    }
}
