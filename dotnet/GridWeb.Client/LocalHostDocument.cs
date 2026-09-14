using System.IO;
namespace GridWeb.Client;

/// <summary>An isolated copy of the trusted embedded host, never workbook-supplied HTML.</summary>
public sealed class LocalHostDocument : IDisposable
{
    public string DirectoryPath { get; }
    public Uri FileUri { get; }
    public const string VirtualOrigin = "https://gridweb.invalid/host.html";
    public LocalHostDocument()
    {
        DirectoryPath = Path.Combine(Path.GetTempPath(), "GridWeb", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(DirectoryPath);
        var file = Path.Combine(DirectoryPath, "host.html");
        var html = HostSession.GetHtml();
        // The offline native host has no legitimate remote resource or connection dependency.
        const string csp = "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; script-src 'unsafe-inline' blob:; style-src 'unsafe-inline'; img-src data: blob:; font-src 'none'; connect-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'\">";
        File.WriteAllText(file, html.Replace("<head>", "<head>" + csp));
        FileUri = new Uri(file);
    }
    public void Dispose()
    {
        try { Directory.Delete(DirectoryPath, true); }
        catch (IOException) { /* WebView may still own a file handle during OS teardown. */ }
        catch (UnauthorizedAccessException) { }
    }
}
