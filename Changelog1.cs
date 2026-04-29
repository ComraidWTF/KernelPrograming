using System.Net;
using System.Text;
using System.Collections.Generic;

public sealed class PropertyChange
{
    public string Name { get; set; } = "";
    public string Before { get; set; } = "";
    public string After { get; set; } = "";
}

public static class EmailHtmlBuilder
{
    public static string BuildChangeSummaryHtml(IEnumerable<PropertyChange> changes)
    {
        var sb = new StringBuilder();

        sb.AppendLine("""
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td align="center">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff">
        
        <tr>
          <td bgcolor="#1f2937" align="center">
            <font face="Arial, sans-serif" color="#ffffff" size="4">
              <b>Change Summary</b>
            </font>
          </td>
        </tr>

        <tr>
          <td>
            <table width="100%" cellpadding="6" cellspacing="0" border="1" bordercolor="#d1d5db">
              
              <tr bgcolor="#f3f4f6">
                <th align="left" width="30%">
                  <font face="Arial, sans-serif" color="#111827">Name</font>
                </th>
                <th align="left" width="35%">
                  <font face="Arial, sans-serif" color="#991b1b">Before</font>
                </th>
                <th align="left" width="35%">
                  <font face="Arial, sans-serif" color="#166534">After</font>
                </th>
              </tr>
""");

        foreach (var change in changes)
        {
            sb.AppendLine($"""
              <tr>
                <td valign="top">
                  <font face="Arial, sans-serif" color="#111827">
                    {EncodeMultiline(change.Name)}
                  </font>
                </td>

                <td valign="top" bgcolor="#fef2f2">
                  <font face="Arial, sans-serif" color="#991b1b">
                    {EncodeMultiline(change.Before)}
                  </font>
                </td>

                <td valign="top" bgcolor="#f0fdf4">
                  <font face="Arial, sans-serif" color="#166534">
                    {EncodeMultiline(change.After)}
                  </font>
                </td>
              </tr>
""");
        }

        sb.AppendLine("""
            </table>
          </td>
        </tr>

        <tr>
          <td align="center">
            <font face="Arial, sans-serif" color="#6b7280" size="2">
              This is an automated notification. Please do not reply.
            </font>
          </td>
        </tr>

      </table>
    </td>
  </tr>
</table>
""");

        return sb.ToString();
    }

    private static string EncodeMultiline(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "-";

        var encoded = WebUtility.HtmlEncode(value)
            .Replace("\r\n", "<br>")
            .Replace("\n", "<br>")
            .Replace("\r", "<br>");

        return InsertBreaks(encoded, 30);
    }

    private static string InsertBreaks(string value, int maxChunkLength)
    {
        var parts = value.Split(' ');

        for (int i = 0; i < parts.Length; i++)
        {
            if (parts[i].Length > maxChunkLength)
            {
                parts[i] = string.Join("<wbr>", Chunk(parts[i], maxChunkLength));
            }
        }

        return string.Join(" ", parts);
    }

    private static IEnumerable<string> Chunk(string value, int size)
    {
        for (int i = 0; i < value.Length; i += size)
        {
            yield return value.Substring(i, System.Math.Min(size, value.Length - i));
        }
    }
}
