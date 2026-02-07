using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.CustomProperties;
using DocxportNet.API;

namespace DocxportNet.Fields;

public sealed class DxpFieldEvalContext
{
	private readonly Dictionary<string, string?> _docVariables = new(StringComparer.OrdinalIgnoreCase);
	private readonly Dictionary<string, string?> _documentProperties = new(StringComparer.OrdinalIgnoreCase);
	private readonly Dictionary<string, DxpFieldValue> _documentPropertyValues = new(StringComparer.OrdinalIgnoreCase);
	private readonly Dictionary<string, string?> _bookmarks = new(StringComparer.OrdinalIgnoreCase);
	private readonly Dictionary<string, string> _mergeFieldAliases = new(StringComparer.OrdinalIgnoreCase);
	private readonly Dictionary<string, int> _sequences = new(StringComparer.OrdinalIgnoreCase);
	private readonly Dictionary<string, string> _numberedItems = new(StringComparer.OrdinalIgnoreCase);
	private readonly Dictionary<string, int> _sequenceResetKeys = new(StringComparer.OrdinalIgnoreCase);

	public Func<DateTimeOffset> NowProvider { get; private set; } = () => DateTimeOffset.Now;
	public CultureInfo? Culture { get; set; } = CultureInfo.CurrentCulture;
	public bool StripAmPmPeriods { get; set; } = false;
	public bool AllowInvariantNumericFallback { get; set; } = true;
	public DocxportNet.Formatting.IDxpNumberToWordsProvider? NumberToWordsProvider { get; set; }
	public DocxportNet.Formatting.DxpNumberToWordsRegistry NumberToWordsRegistry { get; set; } = DocxportNet.Formatting.DxpNumberToWordsRegistry.Default;
	public Resolution.IDxpTableValueResolver? TableResolver { get; set; }
	public Resolution.IDxpRefResolver? RefResolver { get; set; }
	public List<Resolution.DxpRefHyperlink> RefHyperlinks { get; } = new();
	public List<Resolution.DxpRefFootnote> RefFootnotes { get; } = new();
	public Func<int>? CurrentOutlineLevelProvider { get; set; }
	public int? CurrentDocumentOrder { get; set; }
	public Expressions.DxpFormulaFunctionRegistry FormulaFunctions { get; set; } = Expressions.DxpFormulaFunctionRegistry.Default;
	public Resolution.IDxpFieldValueResolver? ValueResolver { get; set; }
	public string? ListSeparator { get; set; }

	public DateTimeOffset? CreatedDate { get; set; }
	public DateTimeOffset? SavedDate { get; set; }
	public DateTimeOffset? PrintDate { get; set; }

	public void SetNow(Func<DateTimeOffset> nowProvider)
	{
		NowProvider = nowProvider ?? throw new ArgumentNullException(nameof(nowProvider));
	}

	public void SetDocVariable(string name, string? value) => _docVariables[name] = value;
	public bool TryGetDocVariable(string name, out string? value) => _docVariables.TryGetValue(name, out value);

	public void SetDocumentProperty(string name, string? value) => _documentProperties[name] = value;
	public bool TryGetDocumentProperty(string name, out string? value) => _documentProperties.TryGetValue(name, out value);
	public void SetDocumentPropertyValue(string name, DxpFieldValue value) => _documentPropertyValues[name] = value;
	public bool TryGetDocumentPropertyValue(string name, out DxpFieldValue value) => _documentPropertyValues.TryGetValue(name, out value);

	public void SetBookmark(string name, string? value) => _bookmarks[name] = value;
	public bool TryGetBookmark(string name, out string? value) => _bookmarks.TryGetValue(name, out value);

	public void SetMergeFieldAlias(string name, string targetName) => _mergeFieldAliases[name] = targetName;
	public bool TryGetMergeFieldAlias(string name, out string? targetName)
		=> _mergeFieldAliases.TryGetValue(name, out targetName);

	public int GetSequence(string identifier) => _sequences.TryGetValue(identifier, out var value) ? value : 0;
	public void SetSequence(string identifier, int value) => _sequences[identifier] = value;
	public int NextSequence(string identifier)
	{
		int next = GetSequence(identifier) + 1;
		_sequences[identifier] = next;
		return next;
	}

	public void SetSequenceResetKey(string key, int value) => _sequenceResetKeys[key] = value;
	public bool TryGetSequenceResetKey(string key, out int value) => _sequenceResetKeys.TryGetValue(key, out value);

	public void SetNumberedItem(string name, string value) => _numberedItems[name] = value;
	public bool TryGetNumberedItem(string name, out string? value)
		=> _numberedItems.TryGetValue(name, out value);

	public Task InitWithDocumentAsync(WordprocessingDocument document, CancellationToken cancellationToken = default)
	{
		return InitWithDocumentAsync(
			document,
			includeBookmarks: false,
			includeDocumentProperties: false,
			includeCustomProperties: false,
			cancellationToken);
	}

	public Task InitWithDocumentAsync(
		WordprocessingDocument document,
		bool includeBookmarks,
		bool includeDocumentProperties = false,
		bool includeCustomProperties = false,
		CancellationToken cancellationToken = default)
	{
		_ = cancellationToken;
		if (document.PackageProperties.Created is DateTime created)
			CreatedDate ??= new DateTimeOffset(created);
		if (document.PackageProperties.Modified is DateTime modified)
			SavedDate ??= new DateTimeOffset(modified);
		if (document.PackageProperties.LastPrinted is DateTime printed)
			PrintDate ??= new DateTimeOffset(printed);

		if (includeDocumentProperties)
			PopulateDocumentProperties(document, includeCustomProperties);

		if (includeBookmarks)
		{
			var bookmarks = Resolution.DxpBookmarkTextExtractor.Extract(document);
			foreach (var kvp in bookmarks)
				SetBookmark(kvp.Key, kvp.Value);
		}

		return Task.CompletedTask;
	}

	private void PopulateDocumentProperties(WordprocessingDocument document, bool includeCustomProperties)
	{
		var core = document.PackageProperties;
		var customList = includeCustomProperties
			? document.CustomFilePropertiesPart?.Properties?
				.Elements<CustomDocumentProperty>()
				.Select(p => new CustomFileProperty(
					p.Name?.Value ?? string.Empty,
					p.FirstChild?.LocalName,
					p.FirstChild?.InnerText))
				.ToList()
			: null;

		PopulateDocumentProperties(core, customList, includeCustomProperties);
	}

	public void InitFromDocumentContext(
		DxpIDocumentContext documentContext,
		bool includeDocumentProperties = false,
		bool includeCustomProperties = false)
	{
		var core = documentContext.CoreProperties;
		if (core != null)
		{
			if (core.Created is DateTime created)
				CreatedDate ??= new DateTimeOffset(created);
			if (core.Modified is DateTime modified)
				SavedDate ??= new DateTimeOffset(modified);
			if (core.LastPrinted is DateTime printed)
				PrintDate ??= new DateTimeOffset(printed);
		}

		if (includeDocumentProperties && core != null)
			PopulateDocumentProperties(core, documentContext.CustomProperties, includeCustomProperties);
	}

	private void PopulateDocumentProperties(
		IPackageProperties core,
		IReadOnlyList<CustomFileProperty>? customProperties,
		bool includeCustomProperties)
	{

		AddDocumentProperty("Title", core.Title);
		AddDocumentProperty("Subject", core.Subject);
		AddDocumentPropertyAliases(core.Creator, "Author", "Creator");
		AddDocumentPropertyAliases(core.Description, "Comments", "Description");
		AddDocumentProperty("Keywords", core.Keywords);
		AddDocumentProperty("Category", core.Category);
		AddDocumentPropertyAliases(core.LastModifiedBy, "LastSavedBy", "Last Saved By", "LastModifiedBy");
		AddDocumentPropertyAliases(core.Revision, "Revision", "Revision Number");
		AddDocumentProperty("Identifier", core.Identifier);
		AddDocumentProperty("ContentStatus", core.ContentStatus);
		AddDocumentProperty("Language", core.Language);
		AddDocumentProperty("Version", core.Version);

		if (core.Created is DateTime created)
			AddDocumentPropertyAliases(new DateTimeOffset(created), "Created", "CreateDate", "Creation Date", "CreationDate");
		if (core.Modified is DateTime modified)
			AddDocumentPropertyAliases(new DateTimeOffset(modified), "Modified", "LastSaved", "Last Saved", "Last Save Time", "LastSavedTime");
		if (core.LastPrinted is DateTime printed)
			AddDocumentPropertyAliases(new DateTimeOffset(printed), "LastPrinted", "Last Printed");

		if (!includeCustomProperties)
			return;

		if (customProperties == null)
			return;

		foreach (var prop in customProperties)
		{
			string? name = prop.Name;
			if (string.IsNullOrWhiteSpace(name))
				continue;
			var value = prop.Value?.ToString();
			AddDocumentProperty(name, value);
		}
	}

	private void AddDocumentProperty(string? name, string? value)
	{
		if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value))
			return;

		var propertyName = name!;
		var propertyValue = value!;
		SetDocumentProperty(propertyName, propertyValue);
		SetDocumentPropertyValue(propertyName, new DxpFieldValue(propertyValue));
	}

	private void AddDocumentPropertyAliases(string? value, params string[] names)
	{
		if (string.IsNullOrWhiteSpace(value))
			return;

		var propertyValue = value!;
		foreach (var name in names)
		{
			SetDocumentProperty(name, propertyValue);
			SetDocumentPropertyValue(name, new DxpFieldValue(propertyValue));
		}
	}

	private void AddDocumentPropertyAliases(DateTimeOffset value, params string[] names)
	{
		foreach (var name in names)
			SetDocumentPropertyValue(name, new DxpFieldValue(value));
	}
}
