using System.Net;
using System.Text;

public static class EmailBodyBuilder
{
    public static string Build(
        Dictionary<string, List<(string Property, string OldValue, string NewValue)>> sections)
    {
        var sb = new StringBuilder();

        sb.AppendLine("""
        <div style="font-family: Aptos, Calibri, Arial, sans-serif; font-size: 14px; color: WindowText; background-color: Window; line-height: 1.5;">

          <p style="margin: 0 0 12px 0;">Hi,</p>

          <p style="margin: 0 0 16px 0;">
            Please find below the list of changes.
          </p>

          <table role="presentation" cellpadding="0" cellspacing="0" border="0"
                 style="border-collapse: collapse; margin-bottom: 20px;">
            <tr>
        """);

        // Top links
        foreach (var section in sections)
        {
            var id = GetId(section.Key);

            sb.AppendLine($$"""
              <td style="padding: 6px 12px 6px 0;">
                <a href="#{{id}}" style="color: LinkText; text-decoration: underline;">
                  {{Html(section.Key)}}
                </a>
              </td>
            """);
        }

        sb.AppendLine("""
            </tr>
          </table>
        """);

        // Sections
        foreach (var section in sections)
        {
            var id = GetId(section.Key);

            sb.AppendLine($$"""
              <h2 id="{{id}}" style="margin: 24px 0 10px 0; font-size: 18px; color: WindowText;">
                {{Html(section.Key)}}
              </h2>

              <table role="presentation"
                     cellpadding="0"
                     cellspacing="0"
                     border="0"
                     style="border-collapse: collapse; width: 100%; max-width: 900px; background-color: Window; color: WindowText; border: 1px solid ButtonBorder; margin-bottom: 20px;">

                <thead>
                  <tr>
                    <th align="left" style="padding: 10px 12px; border-bottom: 1px solid ButtonBorder; background-color: ButtonFace; color: ButtonText;">
                      Property Name
                    </th>
                    <th align="left" style="padding: 10px 12px; border-bottom: 1px solid ButtonBorder; background-color: ButtonFace; color: ButtonText;">
                      Old Value
                    </th>
                    <th align="left" style="padding: 10px 12px; border-bottom: 1px solid ButtonBorder; background-color: ButtonFace; color: ButtonText;">
                      New Value
                    </th>
                  </tr>
                </thead>

                <tbody>
            """);

            if (section.Value.Count == 0)
            {
                sb.AppendLine("""
                  <tr>
                    <td colspan="3" style="padding: 10px 12px;">No changes</td>
                  </tr>
                """);
            }
            else
            {
                foreach (var item in section.Value)
                {
                    sb.AppendLine($$"""
                      <tr>
                        <td style="padding: 10px 12px; border-bottom: 1px solid ButtonBorder; vertical-align: top;">
                          {{Html(item.Property)}}
                        </td>
                        <td style="padding: 10px 12px; border-bottom: 1px solid ButtonBorder; vertical-align: top;">
                          {{Multiline(item.OldValue)}}
                        </td>
                        <td style="padding: 10px 12px; border-bottom: 1px solid ButtonBorder; vertical-align: top; font-weight: 600;">
                          {{Multiline(item.NewValue)}}
                        </td>
                      </tr>
                    """);
                }
            }

            sb.AppendLine("""
                </tbody>
              </table>
            """);
        }

        sb.AppendLine("""
          <p style="margin: 16px 0 0 0;">
            Regards,<br />
            Your Team
          </p>

        </div>
        """);

        return sb.ToString();
    }

    private static string Html(string value) =>
        WebUtility.HtmlEncode(value ?? "");

    private static string Multiline(string value) =>
        string.IsNullOrWhiteSpace(value)
            ? "&nbsp;"
            : Html(value)
                .Replace("\r\n", "<br />")
                .Replace("\n", "<br />")
                .Replace("\r", "<br />");

    private static string GetId(string text) =>
        "sec-" + text.Replace(" ", "-").ToLowerInvariant();
}
