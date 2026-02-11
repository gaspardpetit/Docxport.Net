namespace DocxportNet.Fields.Resolution;

public interface IDxpMergeRecordCursor
{
    bool HasCurrent { get; }
    int RecordIndex { get; }
    bool MoveNext();
    DxpFieldValue? GetValue(string fieldName);
}

public interface IDxpResettableMergeRecordCursor : IDxpMergeRecordCursor
{
    void Reset();
}
