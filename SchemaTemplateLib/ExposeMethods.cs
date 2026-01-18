using SchemaTemplateLib.DataModel;
using SchemaTemplateLib.Interfaces;

namespace SchemaTemplateLib
{
	public class ExposeMethods: IExposeMethods
	{
		private readonly IDataProcessing _dataProcessing;
		public ExposeMethods(IDataProcessing dataProcessing) 
		{
			_dataProcessing = dataProcessing;
		}
		public (MemoryStream Stream, string FileName) GenerateExcelTemplate(string procedureName,Dictionary<string, string> parameterValues)
		{ 
			return _dataProcessing.GenerateExcelTemplate(procedureName, parameterValues);
		}

		public List<ProcedureParam> GetProcedureParams(string procedureName)
		{
			return _dataProcessing.GetProcedureParams(procedureName);
		}
	}
}
