using System;
using System.Text.RegularExpressions;

public class MyData
{
    public string TypeA { get; set; }
    public string TypeB { get; set; }
    public string TypeC { get; set; }
    public string TypeD { get; set; }
}

public static class Parser
{
    public static MyData Parse(string input)
    {
        var result = new MyData();

        var regex = new Regex(
            @"(?ms)^(TypeA|TypeB|TypeC|TypeD):\s*(.*?)(?=^TypeA:|^TypeB:|^TypeC:|^TypeD:|\z)"
        );

        foreach (Match match in regex.Matches(input))
        {
            var key = match.Groups[1].Value;
            var value = match.Groups[2].Value.Trim();

            switch (key)
            {
                case "TypeA":
                    result.TypeA = value;
                    break;

                case "TypeB":
                    result.TypeB = value;
                    break;

                case "TypeC":
                    result.TypeC = value;
                    break;

                case "TypeD":
                    result.TypeD = value;
                    break;
            }
        }

        return result;
    }
}
