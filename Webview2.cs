using Microsoft.Web.WebView2.Wpf;
using System.Windows;

namespace YourAppNamespace;

public static class WebView2Html
{
    public static readonly DependencyProperty HtmlProperty =
        DependencyProperty.RegisterAttached(
            "Html",
            typeof(string),
            typeof(WebView2Html),
            new PropertyMetadata(null, OnHtmlChanged));

    public static string GetHtml(DependencyObject obj)
    {
        return (string)obj.GetValue(HtmlProperty);
    }

    public static void SetHtml(DependencyObject obj, string value)
    {
        obj.SetValue(HtmlProperty, value);
    }

    private static async void OnHtmlChanged(
        DependencyObject d,
        DependencyPropertyChangedEventArgs e)
    {
        if (d is not WebView2 webView)
            return;

        var html = e.NewValue as string;

        if (string.IsNullOrWhiteSpace(html))
            return;

        await webView.EnsureCoreWebView2Async();

        webView.NavigateToString(html);
    }
}
