using System.Net;
using System.Text;

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
            <table width="100%" cellpadding="5" cellspacing="0" border="1" bordercolor="#d1d5db">
              <tr bgcolor="#f3f4f6">
                <th align="left" width="25%">
                  <font face="Arial, sans-serif" color="#111827" size="2">Name</font>
                </th>
                <th align="left" width="37%">
                  <font face="Arial, sans-serif" color="#991b1b" size="2">Before</font>
                </th>
                <th align="left" width="38%">
                  <font face="Arial, sans-serif" color="#166534" size="2">After</font>
                </th>
              </tr>
""");

        foreach (var change in changes)
        {
            sb.AppendLine($"""
              <tr>
                <td valign="top">
                  <font face="Arial, sans-serif" color="#111827" size="2">
                    {EncodeValue(change.Name)}
                  </font>
                </td>

                <td valign="top" bgcolor="#fef2f2">
                  <font face="Arial, sans-serif" color="#991b1b" size="2">
                    {EncodeValue(change.Before)}
                  </font>
                </td>

                <td valign="top" bgcolor="#f0fdf4">
                  <font face="Arial, sans-serif" color="#166534" size="2">
                    {EncodeValue(change.After)}
                  </font>
                </td>
              </tr>
""");
        }

        sb.AppendLine("""
            </table>
          </td>
        </tr>

      </table>
    </td>
  </tr>
</table>
""");

        return sb.ToString();
    }

    private static string EncodeValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "-";

        return WebUtility.HtmlEncode(value)
            .Replace("\r\n", "<br>")
            .Replace("\n", "<br>")
            .Replace("\r", "<br>");
    }
}
