public sealed class TokenFileStore
{
    private readonly string _filePath;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public TokenFileStore(string filePath)
    {
        _filePath = filePath;
    }

    public async Task WriteTokenAsync(
        string token,
        CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken);

        try
        {
            string? directory = Path.GetDirectoryName(_filePath);

            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            await RetryAsync(async () =>
            {
                await File.WriteAllTextAsync(
                    _filePath,
                    token,
                    cancellationToken);
            }, cancellationToken);
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<string?> ReadTokenAsync(
        CancellationToken cancellationToken = default)
    {
        await _lock.WaitAsync(cancellationToken);

        try
        {
            if (!File.Exists(_filePath))
            {
                return null;
            }

            return await RetryAsync(
                () => File.ReadAllTextAsync(_filePath, cancellationToken),
                cancellationToken);
        }
        finally
        {
            _lock.Release();
        }
    }

    private static async Task RetryAsync(
        Func<Task> operation,
        CancellationToken cancellationToken)
    {
        await RetryAsync(
            async () =>
            {
                await operation();
                return true;
            },
            cancellationToken);
    }

    private static async Task<T> RetryAsync<T>(
        Func<Task<T>> operation,
        CancellationToken cancellationToken)
    {
        const int maximumAttempts = 5;

        for (int attempt = 1; attempt <= maximumAttempts; attempt++)
        {
            try
            {
                return await operation();
            }
            catch (IOException) when (attempt < maximumAttempts)
            {
                await Task.Delay(
                    TimeSpan.FromMilliseconds(100 * attempt),
                    cancellationToken);
            }
        }

        throw new InvalidOperationException("Unreachable retry state.");
    }
}