namespace SchemaTemplateLib.DataModel
{
	public class ProcedureParam
	{
		public string Name { get; set; }
		public string DataType { get; set; }
		public bool HasDefaultValue { get; set; }
		public object DefaultValue { get; set; }
	}
}
