namespace Excel2MessagePack
{
    public class Def
    {
        public static Type GetDotNetType(string value) => value switch
        {
            "int" => typeof(int),
            "doub" => typeof(double),
            "bool" => typeof(bool),
            "decimal" => typeof(decimal),
            "char" => typeof(char),
            "long" => typeof(long),
            "short" => typeof(short),
            "byte" => typeof(byte),
            "sbyte" => typeof(sbyte),
            "float" => typeof(float),
            "uint" => typeof(uint),
            "ulong" => typeof(ulong),
            "ushort" => typeof(ushort),
            _ => typeof(string)
        };
    }
}
