This is the real gotcha. When an item is **outside the viewport**, virtualization means its container simply doesn't exist yet — `ContainerFromIndex` returns `null` even after `UpdateLayout()`. You need to **force the virtualizing panel to scroll the index into view first**, which triggers container generation.

## The Fix — Force the VirtualizingPanel

```csharp
public static async Task SelectAndRevealByIndexPathAsync(
    RadTreeView treeView,
    int[] indexPath)
{
    ItemsControl currentParent = treeView;
    RadTreeViewItem container = null;

    for (int depth = 0; depth < indexPath.Length; depth++)
    {
        int index = indexPath[depth];

        await Application.Current.Dispatcher.InvokeAsync(
            () => { }, DispatcherPriority.Background);

        // 👇 Force the virtualizing panel to generate the container
        await BringIndexIntoViewAsync(currentParent, index);

        container = currentParent.ItemContainerGenerator
                                 .ContainerFromIndex(index) as RadTreeViewItem;

        if (container == null)
        {
            Debug.WriteLine($"Failed at depth {depth}, index {index}");
            return;
        }

        bool isTarget = depth == indexPath.Length - 1;

        if (isTarget)
        {
            container.IsSelected = true;
            await Application.Current.Dispatcher.InvokeAsync(
                () => container.BringIntoView(), DispatcherPriority.Loaded);
        }
        else
        {
            if (!container.IsExpanded)
            {
                container.IsExpanded = true;
                await Application.Current.Dispatcher.InvokeAsync(
                    () => container.UpdateLayout(), DispatcherPriority.Loaded);
            }

            currentParent = container;
        }
    }
}
```

---

## The Key Method — BringIndexIntoViewAsync

```csharp
private static async Task BringIndexIntoViewAsync(ItemsControl itemsControl, int index)
{
    await Application.Current.Dispatcher.InvokeAsync(() =>
    {
        var panel = GetVirtualizingStackPanel(itemsControl);

        if (panel != null)
        {
            // BringIndexIntoViewPublic forces the VSP to generate
            // the container even if it's off-screen
            var method = typeof(VirtualizingStackPanel).GetMethod(
                "BringIndexIntoViewPublic",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            method?.Invoke(panel, new object[] { index });
        }

        itemsControl.UpdateLayout();

    }, DispatcherPriority.Loaded);

    // One more yield so the layout pass fully completes
    await Application.Current.Dispatcher.InvokeAsync(
        () => { }, DispatcherPriority.Background);
}
```

---

## VirtualizingStackPanel Finder

```csharp
private static VirtualizingStackPanel GetVirtualizingStackPanel(ItemsControl itemsControl)
{
    var itemsHost = itemsControl.ItemContainerGenerator
                                .ContainerFromIndex(0); // nudge generator

    return FindVisualChild<VirtualizingStackPanel>(itemsControl);
}

private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
{
    for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
    {
        var child = VisualTreeHelper.GetChild(parent, i);

        if (child is T match)
            return match;

        var result = FindVisualChild<T>(child);
        if (result != null)
            return result;
    }

    return null;
}
```

---

## What's Happening Under the Hood

```
Without fix:                         With fix:
─────────────────────────────────    ──────────────────────────────────────
ContainerFromIndex(5)                BringIndexIntoViewPublic(5)
  → null (item off-screen,             → VSP scrolls internally to index 5
    container never generated)         → WPF generates the container
                                       → layout pass completes
                                     ContainerFromIndex(5)
                                       → ✅ RadTreeViewItem returned
```

| Scenario | Handled |
|---|---|
| Item above viewport | ✅ VSP scrolls up, generates container |
| Item below viewport | ✅ VSP scrolls down, generates container |
| Deeply nested & off-screen | ✅ Each level forced into view before descending |
| Already in viewport | ✅ No-op, works as before |

`BringIndexIntoViewPublic` is the internal WPF method that `ListBox.ScrollIntoView` calls under the hood — it's the right tool for the job here even though it requires reflection.
