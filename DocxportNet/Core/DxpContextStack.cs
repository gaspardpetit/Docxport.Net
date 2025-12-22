namespace DocxportNet.Core;

internal sealed class DxpContextStack<T> where T : class
{
	private readonly List<T> _items = new();
	private readonly string _name;

	public DxpContextStack(string name = "context")
	{
		_name = name;
	}

	public T? Current => _items.Count == 0 ? null : _items[_items.Count - 1];

	public int Count => _items.Count;

	public IDisposable Push(T item)
	{
		_items.Add(item);
		return DxpDisposable.Create(() => Pop(item));
	}

	private void Pop(T expected)
	{
		if (_items.Count == 0 || !ReferenceEquals(expected, _items[_items.Count - 1]))
			throw new InvalidOperationException($"Attempted to pop a {_name} that was not the last");
		_items.RemoveAt(_items.Count - 1);
	}
}
