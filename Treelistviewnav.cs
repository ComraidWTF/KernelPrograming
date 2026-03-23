private void NavigateToPath(RadTreeListView treeListView, IList<TreeNode> rootCollection, int[] path)
{
    if (path == null || path.Length == 0) return;

    TreeNode current = null;
    IList<TreeNode> currentLevel = rootCollection;

    foreach (int index in path)
    {
        if (index < 0 || index >= currentLevel.Count) return; // Guard: invalid index

        current = currentLevel[index];

        // Expand this node so its children are realized in the UI
        treeListView.ExpandHierarchyItem(current);

        // Move down to the next level
        currentLevel = current.Children;
    }

    if (current == null) return;

    // Select the final item
    treeListView.SelectedItem = current;

    // Scroll + focus after layout pass (virtualization needs time to render)
    treeListView.Dispatcher.BeginInvoke(new Action(() =>
    {
        treeListView.ScrollIntoView(current);
        treeListView.UpdateLayout();

        var row = treeListView.ItemContainerGenerator
                              .ContainerFromItem(current) as GridViewRow;
        row?.Focus();

    }), System.Windows.Threading.DispatcherPriority.Background);
}
