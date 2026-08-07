public static void CleanupOldLogs(string logRoot, int retentionDays = 60)
{
    if (string.IsNullOrWhiteSpace(logRoot))
        return;

    string root = Path.GetFullPath(logRoot)
        .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

    if (!Directory.Exists(root))
        return;

    // Safety: never allow filesystem root like C:\ or /
    string? driveRoot = Path.GetPathRoot(root)?
        .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

    if (string.Equals(
        root,
        driveRoot,
        OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal))
    {
        return;
    }

    DateTime cutoff = DateTime.UtcNow.AddDays(-retentionDays);

    // Delete old .ulog files
    foreach (string file in Directory.EnumerateFiles(
                 root,
                 "*.ulog",
                 SearchOption.AllDirectories))
    {
        try
        {
            string fullPath = Path.GetFullPath(file);

            string rootPrefix = root + Path.DirectorySeparatorChar;

            bool isInsideRoot = fullPath.StartsWith(
                rootPrefix,
                OperatingSystem.IsWindows()
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal);

            if (!isInsideRoot)
                continue;

            if (File.GetLastWriteTimeUtc(fullPath) < cutoff)
                File.Delete(fullPath);
        }
        catch (IOException)
        {
            // File may be in use - ignore and continue
        }
        catch (UnauthorizedAccessException)
        {
            // No permission - ignore and continue
        }
    }

    // Delete empty subfolders, deepest first.
    // Never deletes logRoot itself.
    foreach (string directory in Directory
                 .EnumerateDirectories(root, "*", SearchOption.AllDirectories)
                 .OrderByDescending(x => x.Length))
    {
        try
        {
            string fullPath = Path.GetFullPath(directory);

            string rootPrefix = root + Path.DirectorySeparatorChar;

            bool isInsideRoot = fullPath.StartsWith(
                rootPrefix,
                OperatingSystem.IsWindows()
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal);

            if (!isInsideRoot)
                continue;

            var info = new DirectoryInfo(fullPath);

            // Don't touch junctions/symlinks
            if ((info.Attributes & FileAttributes.ReparsePoint) != 0)
                continue;

            if (!Directory.EnumerateFileSystemEntries(fullPath).Any())
                Directory.Delete(fullPath);
        }
        catch (IOException)
        {
            // Folder may be in use or no longer empty
        }
        catch (UnauthorizedAccessException)
        {
            // Ignore and continue
        }
    }
}