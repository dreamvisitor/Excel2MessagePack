using System.Reflection;

namespace Excel2MessagePack
{
    public class Setting
    {
        public string TargetFolder { get; set; } = string.Empty;
        public string OutputFolder { get; set; } = string.Empty;
        public string SourceCodeFolder { get; set; } = string.Empty;
    }

    internal class ReadIni
    {
        public static Setting? ReadSettings(string iniFilePath)
        {
            var settings = new Setting();

            Dictionary<string, string>? iniData = ReadIniFile(iniFilePath);

            if (iniData == null)
                return null;

            // 3. 将 INI 文件的键值对赋值给 Setting 类的同名属性
            foreach (var kvp in iniData)
            {
                PropertyInfo? property = typeof(Setting).GetProperty(kvp.Key, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);
                if (property != null && property.CanWrite)
                {
                    // 将字符串值转换为属性的类型
                    var value = Convert.ChangeType(kvp.Value, property.PropertyType);
                    property.SetValue(settings, value);
                }
            }

            return settings;
        }

        // 读取 INI 文件并解析为字典
        static Dictionary<string, string>? ReadIniFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                using StreamWriter sw = new(filePath);
                sw.WriteLine("[Settings]");
                sw.WriteLine($"TargetFolder=./Out/TargetFolder");
                sw.WriteLine($"OutputFolder=./Out/OutputFolder");
                sw.WriteLine($"SourceCodeFolder=./Out/SourceCodeFolder");
                sw.Flush();
                sw.Close();

                return null;
            }
            else
            {
                var iniData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (var line in File.ReadAllLines(filePath))
                {
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith(";") || line.StartsWith("#"))
                        continue; // 忽略空行和注释

                    var parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        string key = parts[0].Trim();
                        string value = parts[1].Trim();
                        iniData[key] = value;
                    }
                }

                return iniData;
            }
        }
    }
}
