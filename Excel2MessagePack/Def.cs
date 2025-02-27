namespace Excel2MessagePack
{
    public class Def
    {
        public static Type GetDotNetType(string value) => value switch
        {
            "int" => typeof(int),
            "double" => typeof(double),
            "bool" => typeof(bool),
            "long" => typeof(long),
            "short" => typeof(short),
            "float" => typeof(float),
            _ => typeof(string)
        };
    }
}
