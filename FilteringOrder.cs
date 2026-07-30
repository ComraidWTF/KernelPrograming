public static void ApplyFilteredOrderById<T, TId>(
    IList<T> source,
    IReadOnlyList<T> reorderedFilteredItems,
    Func<T, bool> filter,
    Func<T, TId> idSelector)
    where TId : notnull
{
    var reorderedIds = reorderedFilteredItems
        .Select(idSelector)
        .ToArray();

    var filteredSourceIds = source
        .Where(filter)
        .Select(idSelector)
        .ToHashSet();

    if (reorderedIds.Length != filteredSourceIds.Count ||
        reorderedIds.Any(id => !filteredSourceIds.Contains(id)))
    {
        throw new InvalidOperationException(
            "The reordered items do not match the filtered source items.");
    }

    var itemById = source.ToDictionary(idSelector);
    var reorderedIndex = 0;

    for (var index = 0; index < source.Count; index++)
    {
        if (!filter(source[index]))
        {
            continue;
        }

        source[index] = itemById[reorderedIds[reorderedIndex]];
        reorderedIndex++;
    }
}