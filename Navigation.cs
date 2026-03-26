public class TaskNavigator
{
    private readonly Stack<int> _backStack = new();
    private readonly Stack<int> _forwardStack = new();
    private int? _current;
    private bool _isInternalNavigation; // ← flag

    public int? Current => _current;
    public bool CanGoBack => _backStack.Count > 0;
    public bool CanGoForward => _forwardStack.Count > 0;

    public void NavigateTo(int taskId)
    {
        if (!_isInternalNavigation)
        {
            if (_current.HasValue)
                _backStack.Push(_current.Value);
            _forwardStack.Clear();
        }

        _current = taskId;
        // ... rest of your logic
    }

    public int? GoBack()
    {
        if (!CanGoBack) return null;

        _forwardStack.Push(_current!.Value);

        _isInternalNavigation = true;
        NavigateTo(_backStack.Pop());
        _isInternalNavigation = false;

        return _current;
    }

    public int? GoForward()
    {
        if (!CanGoForward) return null;

        _backStack.Push(_current!.Value);

        _isInternalNavigation = true;
        NavigateTo(_forwardStack.Pop());
        _isInternalNavigation = false;

        return _current;
    }
}
