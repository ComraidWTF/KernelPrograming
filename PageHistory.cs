using System.Collections.Generic;

public class TaskNavigationManager
{
    private readonly Stack<int> _backStack = new();
    private readonly Stack<int> _forwardStack = new();

    private bool _suppressHistoryUpdate;

    public int? CurrentTaskId { get; private set; }

    public bool CanGoBack => _backStack.Count > 0;
    public bool CanGoForward => _forwardStack.Count > 0;

    public void SelectTask(int taskId)
    {
        if (CurrentTaskId == taskId)
            return;

        if (!_suppressHistoryUpdate)
        {
            if (CurrentTaskId.HasValue)
                _backStack.Push(CurrentTaskId.Value);

            _forwardStack.Clear();
        }

        CurrentTaskId = taskId;
    }

    public int? GoBack()
    {
        if (!CanGoBack)
            return CurrentTaskId;

        if (CurrentTaskId.HasValue)
            _forwardStack.Push(CurrentTaskId.Value);

        var previousTaskId = _backStack.Pop();

        _suppressHistoryUpdate = true;
        try
        {
            SelectTask(previousTaskId);
        }
        finally
        {
            _suppressHistoryUpdate = false;
        }

        return CurrentTaskId;
    }

    public int? GoForward()
    {
        if (!CanGoForward)
            return CurrentTaskId;

        if (CurrentTaskId.HasValue)
            _backStack.Push(CurrentTaskId.Value);

        var nextTaskId = _forwardStack.Pop();

        _suppressHistoryUpdate = true;
        try
        {
            SelectTask(nextTaskId);
        }
        finally
        {
            _suppressHistoryUpdate = false;
        }

        return CurrentTaskId;
    }
}
