using MessagePack;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using OfficeOpenXml;
using System.Reflection;
using System.Text;

namespace Excel2MessagePack
{
    internal class ProcessExcel(ExcelWorksheet worksheet, Setting settings)
    {
        public void Process()
        {
            Console.WriteLine($"开始处理工作表: {worksheet.Name} 列：{worksheet.Dimension.End.Row} 行：{worksheet.Dimension.End.Column}");

            // 获取工作表名称
            string className = worksheet.Name;

            List<(string Name, string Type, string description)> properties = [];
            Dictionary<int, string> propertyKey = [];

            // 获取工作表的行数和列数
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // 容器类型
            string containerType = worksheet.Cells[1, 1].Text;

            // 遍历工作表的前三行 获取属性名 类型 描述
            for (int col = 1; col <= colCount; col++)
            {
                string propertyName = worksheet.Cells[2, col].Text;
                string propertyType = worksheet.Cells[3, col].Text;
                string description = worksheet.Cells[4, col].Text;
                properties.Add((propertyName, propertyType, description));
                propertyKey.Add(col, propertyName);
            }

            // 创建类文件
            string code = GenerateClassFile(settings.SourceCodeFolder, className, properties);

            // 创建动态类
            Assembly? assembly = DynamicClass(code);
            if (assembly == null)
            {
                Console.WriteLine("无法创建动态类");
                return;
            }

            Type? type = assembly.GetType(className);
            if (type == null)
                return;

            byte[]? bytes = null;
            if (containerType.Equals("list", StringComparison.OrdinalIgnoreCase))
            {
                bytes = CreateMessagePackDataList(type, rowCount, colCount, className, propertyKey);
            }
            else
            {
                bytes = CreateMessagePackDataMap(properties[0].Type, type, rowCount, colCount, className, propertyKey);
            }

            if (bytes != null)
            {
                // 将二进制数据保存到文件
                string outputFilePath = Path.Combine(settings.OutputFolder, $"{className}.bytes");
                File.WriteAllBytes(outputFilePath, bytes);

                // 将字节数组转换为 JSON
                string json = MessagePackSerializer.ConvertToJson(bytes);

                // 将 JSON 写入文件
                string jsonFilePath = Path.Combine(settings.OutputFolder, $"{className}.json");
                File.WriteAllText(jsonFilePath, json);
            }
        }


        byte[] CreateMessagePackDataList(Type valueType, int rowCount, int colCount, string className, Dictionary<int, string> propertyKey)
        {
            // 创建List
            var listType = typeof(List<>).MakeGenericType(valueType);
            Console.WriteLine("动态创建 列表 Type: " + listType);

            var list = Activator.CreateInstance(listType);

            Console.WriteLine("开始填充数据");

            // 填充数据
            for (int row = 5; row <= rowCount; row++)
            {
                object? obj = Activator.CreateInstance(valueType);
                if (obj == null)
                {
                    Console.WriteLine($"无法创建类的实例: {className}");
                    return [];
                }

                // 设置元素
                for (int col = 1; col <= colCount; col++)
                {
                    string otherData = worksheet.Cells[row, col].Text;

                    // 获取属性信息
                    PropertyInfo? otherInfo = valueType.GetProperty(propertyKey[col]);
                    if (otherInfo == null)
                    {
                        Console.WriteLine($"无法找到属性: {propertyKey[col]}");
                        continue;
                    }

                    // 设置属性值
                    Type otherType = otherInfo.PropertyType;
                    object? otherValue = Convert.ChangeType(otherData, otherType);
                    otherInfo.SetValue(obj, otherValue);
                }

                // 调用字典的 Add 方法
                listType.GetMethod("Add")?.Invoke(list, [obj]);
            }

            // 2. 注册动态解析器
            var resolver = MessagePack.Resolvers.CompositeResolver.Create(
                MessagePack.Resolvers.DynamicObjectResolver.Instance,
                MessagePack.Resolvers.StandardResolver.Instance
            );

            var options = MessagePackSerializerOptions.Standard.WithResolver(resolver);

            // 将字典转换为字节数组
            byte[] bytes = MessagePackSerializer.Serialize(list, options);

            return bytes;
        }

        byte[] CreateMessagePackDataMap(string keyTypeStr, Type valueType, int rowCount, int colCount, string className, Dictionary<int, string> propertyKey)
        {
            // 创建字典
            var keyType = Def.GetDotNetType(keyTypeStr);
            if (keyType == null)
                return [];

            var dictType = typeof(Dictionary<,>).MakeGenericType(keyType, valueType);
            Console.WriteLine("动态创建 字典 Type: " + dictType);

            var dict = Activator.CreateInstance(dictType);

            for (int row = 5; row <= rowCount; row++)
            {
                object? obj = Activator.CreateInstance(valueType);
                if (obj == null)
                {
                    Console.WriteLine($"无法创建类的实例: {className}");
                    return [];
                }

                // 设置 第一个元素 为 key
                string keyData = worksheet.Cells[row, 1].Text;

                PropertyInfo? keyInfo = valueType.GetProperty(propertyKey[1]);
                if (keyInfo == null)
                {
                    Console.WriteLine($"无法找到属性: {propertyKey[1]}");
                    continue;
                }

                // 设置key属性值
                object? convertedValue = Convert.ChangeType(keyData, keyInfo.PropertyType);
                keyInfo.SetValue(obj, convertedValue);

                // 设置其他元素
                for (int col = 2; col <= colCount; col++)
                {
                    string otherData = worksheet.Cells[row, col].Text;

                    // 获取属性信息
                    PropertyInfo? otherInfo = valueType.GetProperty(propertyKey[col]);
                    if (otherInfo == null)
                    {
                        Console.WriteLine($"无法找到属性: {propertyKey[col]}");
                        continue;
                    }

                    // 设置属性值
                    Type otherType = otherInfo.PropertyType;
                    object? otherValue = Convert.ChangeType(otherData, otherType);
                    otherInfo.SetValue(obj, otherValue);
                }

                // 调用字典的 Add 方法
                dictType.GetMethod("Add")?.Invoke(dict, [convertedValue, obj]);
            }

            // 2. 注册动态解析器
            var resolver = MessagePack.Resolvers.CompositeResolver.Create(
                MessagePack.Resolvers.DynamicObjectResolver.Instance,
                MessagePack.Resolvers.StandardResolver.Instance
            );

            var options = MessagePackSerializerOptions.Standard.WithResolver(resolver);

            // 将字典转换为字节数组
            byte[] bytes = MessagePackSerializer.Serialize(dict, options);

            return bytes;
        }

        static string GenerateClassFile(string sourceCodeFolder, string className, List<(string Name, string Type, string description)> properties)
        {
            var sb = new StringBuilder();
            sb.AppendLine("// ================================================");
            sb.AppendLine($"// MessagePack Object");
            sb.AppendLine($"// {className}");
            sb.AppendLine($"// {DateTime.Now}");
            sb.AppendLine($"// Created by Excel2MessagePack");
            sb.AppendLine("// ================================================");
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine("using System;");
            sb.AppendLine("using System.Collections.Generic;");
            sb.AppendLine("using MessagePack;");
            sb.AppendLine();
            sb.AppendLine("[MessagePackObject]");
            sb.AppendLine($"public class {className}");
            sb.AppendLine("{");

            for (int i = 0; i < properties.Count; i++)
            {
                var property = properties[i];
                if (i != 0)
                {
                    sb.AppendLine();
                }

                sb.AppendLine($"    // {property.description}");
                sb.AppendLine($"    [Key({i})]");
                sb.AppendLine($"    public {property.Type} {property.Name} {{ get; set; }}");
            }

            sb.AppendLine("}");

            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine("// ================================================");
            sb.AppendLine("// {className}Mgr");
            sb.AppendLine($"public class {className}Mgr");
            sb.AppendLine("{");
            sb.AppendLine($"    public static Dictionary<{properties[0].Type}, {className}> CfgDict;");
            sb.AppendLine("}");


            string code = sb.ToString();
            string filePath = Path.Combine(sourceCodeFolder, $"{className}.cs");
            File.WriteAllText(filePath, code);

            return code;
        }

        static Assembly? DynamicClass(string code)
        {
            // 2. 使用 Roslyn 编译代码
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(code);

            // 添加必要的元数据引用
            var references = new MetadataReference[]
            {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(MessagePackObjectAttribute).Assembly.Location),
                MetadataReference.CreateFromFile(Path.Combine(Path.GetDirectoryName(typeof(object).Assembly.Location), "System.Runtime.dll")),
                MetadataReference.CreateFromFile(Path.Combine(Path.GetDirectoryName(typeof(object).Assembly.Location), "netstandard.dll")),
            };

            // 创建编译选项
            var compilation = CSharpCompilation.Create(
                "DynamicAssembly",
                [syntaxTree],
                references,
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, optimizationLevel: OptimizationLevel.Debug));

            // 编译代码到内存流
            using var ms = new MemoryStream();
            EmitResult result = compilation.Emit(ms);

            // 检查编译是否成功
            if (!result.Success)
            {
                foreach (var diagnostic in result.Diagnostics)
                {
                    Console.WriteLine("编译: " + diagnostic.ToString());
                }
                return null;
            }

            // 3. 加载程序集
            ms.Seek(0, SeekOrigin.Begin);
            Assembly? assembly = Assembly.Load(ms.ToArray());

            return assembly;
        }
    }
}
