

using SchemaTemplateLib.DataModel;

namespace SchemaTemplateLib.Interfaces
{
	public interface IDataProcessing
	{
		(MemoryStream Stream, string FileName) GenerateExcelTemplate(string procedureName, Dictionary<string, string> parameterValues);
		List<ProcedureParam> GetProcedureParams(string procedureName);
		List<string> SearchProcedures(string query);
	}
}
