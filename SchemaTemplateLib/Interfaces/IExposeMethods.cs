using System;
using System.Collections.Generic;
using System.Text;

using System.Text;
using SchemaTemplateLib.DataModel;

namespace SchemaTemplateLib.Interfaces
{
	public interface IExposeMethods
	{
		(MemoryStream Stream, string FileName) GenerateExcelTemplate(string procedureName, Dictionary<string, string> parameterValues);
		List<ProcedureParam> GetProcedureParams(string procedureName);
	}
}
