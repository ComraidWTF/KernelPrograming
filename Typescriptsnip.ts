function normalizeTree(nodes: RawNode[], parentPath = ''): TreeNode[] {
  return nodes.map((node, index) => {
    const uiId = `${parentPath}/${node.type}-${node.id}-${index}`;

    return {
      ...node,
      uiId,
      originalId: node.id,
      children: node.children
        ? normalizeTree(node.children, uiId)
        : undefined,
    };
  });
}
