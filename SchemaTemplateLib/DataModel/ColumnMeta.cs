
namespace SchemaTemplateLib.DataModel;

public class ColumnMeta
{
	public string ColumnKey { get; set; }
	public string DisplayName { get; set; }
	public char DataType { get; set; }
	public char Alignment { get; set; }
	public string Header { get; set; }
	public int ExcelColIndex { get; set; }
}
