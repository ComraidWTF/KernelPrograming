public static void CleanupOldLogs(string logRoot, int retentionDays = 60)
{
    if (string.IsNullOrWhiteSpace(logRoot) || !Directory.Exists(logRoot))
        return;

    var cutoff = DateTime.UtcNow.AddDays(-retentionDays);

    // Delete old .ulog files
    foreach (var file in Directory.EnumerateFiles(
                 logRoot,
                 "*.ulog",
                 SearchOption.AllDirectories))
    {
        try
        {
            if (File.GetLastWriteTimeUtc(file) < cutoff)
                File.Delete(file);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    // Delete empty folders, deepest first
    foreach (var directory in Directory
                 .EnumerateDirectories(logRoot, "*", SearchOption.AllDirectories)
                 .OrderByDescending(x => x.Length))
    {
        try
        {
            if (!Directory.EnumerateFileSystemEntries(directory).Any())
                Directory.Delete(directory);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}