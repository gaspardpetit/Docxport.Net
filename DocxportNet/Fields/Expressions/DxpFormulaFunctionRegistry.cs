namespace DocxportNet.Fields.Expressions;

public delegate double DxpFormulaFunction(IReadOnlyList<double> args);

public sealed class DxpFormulaFunctionRegistry
{
	private readonly Dictionary<string, DxpFormulaFunction> _functions = new(StringComparer.OrdinalIgnoreCase);

	public static DxpFormulaFunctionRegistry Default { get; } = CreateDefault();

	public void Register(string name, DxpFormulaFunction fn)
	{
		if (string.IsNullOrWhiteSpace(name))
			throw new ArgumentException("Function name is required.", nameof(name));
		if (fn == null)
			throw new ArgumentNullException(nameof(fn));
		_functions[name] = fn;
	}

	public bool TryResolve(string name, out DxpFormulaFunction fn) => _functions.TryGetValue(name, out fn!);

	private static DxpFormulaFunctionRegistry CreateDefault()
	{
		var registry = new DxpFormulaFunctionRegistry();
		registry.Register("SUM", args => args.Sum());
		registry.Register("PRODUCT", args => args.Aggregate(1.0, (a, b) => a * b));
		registry.Register("AVERAGE", args => args.Count == 0 ? 0 : args.Sum() / args.Count);
		registry.Register("COUNT", args => args.Count);
		registry.Register("MIN", args => args.Count == 0 ? 0 : args.Min());
		registry.Register("MAX", args => args.Count == 0 ? 0 : args.Max());
		registry.Register("ABS", args => args.Count > 0 ? Math.Abs(args[0]) : 0);
		registry.Register("INT", args => args.Count > 0 ? Math.Floor(args[0]) : 0);
		registry.Register("MOD", args => args.Count >= 2 && args[1] != 0 ? args[0] % args[1] : 0);
		registry.Register("NOT", args => args.Count > 0 && args[0] != 0 ? 0 : 1);
		registry.Register("AND", args => args.All(a => a != 0) ? 1 : 0);
		registry.Register("OR", args => args.Any(a => a != 0) ? 1 : 0);
		registry.Register("TRUE", args => 1);
		registry.Register("FALSE", args => 0);
		registry.Register("ROUND", args =>
			args.Count >= 2 ? Math.Round(args[0], (int)args[1], MidpointRounding.AwayFromZero) :
			args.Count == 1 ? Math.Round(args[0], 0, MidpointRounding.AwayFromZero) : 0);
		registry.Register("SIGN", args => args.Count > 0 ? Math.Sign(args[0]) : 0);
		return registry;
	}
}
