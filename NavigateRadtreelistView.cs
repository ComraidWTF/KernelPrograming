private async void NavigateToPath(RadTreeListView treeListView, IList yourRootCollection, int[] path)
{
    if (path == null || path.Length == 0) return;

    object current = null;
    IList currentLevel = yourRootCollection;

    foreach (int index in path)
    {
        if (index < 0 || index >= currentLevel.Count) return;

        current = currentLevel[index];

        treeListView.ExpandHierarchyItem(current);

        await FlushDispatcher();

        var childrenProp = current.GetType().GetProperty("YourChildrenPropertyName");
        currentLevel = childrenProp?.GetValue(current) as IList;

        if (currentLevel == null) break;
    }

    if (current == null) return;

    treeListView.SelectedItem = current;
    treeListView.ScrollIntoView(current);

    await FlushDispatcher();
    treeListView.UpdateLayout();

    var row = await WaitForContainer(treeListView, current);
    row?.Focus();
}

private async Task<GridViewRow> WaitForContainer(RadTreeListView tree, object item, int maxRetries = 5)
{
    for (int i = 0; i < maxRetries; i++)
    {
        var row = tree.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
        if (row != null) return row;

        await FlushDispatcher();
        tree.UpdateLayout();
    }

    return null;
}

private Task FlushDispatcher()
{
    var tcs = new TaskCompletionSource<bool>();
    Application.Current.Dispatcher.BeginInvoke(
        new Action(() => tcs.SetResult(true)),
        DispatcherPriority.Loaded
    );
    return tcs.Task;
}
