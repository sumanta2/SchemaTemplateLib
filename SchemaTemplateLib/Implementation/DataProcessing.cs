using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using SchemaTemplateLib.DataModel;
using SchemaTemplateLib.Interfaces;
using System.Data;

namespace SchemaTemplateLib.Implementation;

public class DataProcessing : IDataProcessing
{
	private static bool param = false;
	private static readonly HashSet<string> ValidMarkers = new HashSet<string>
	{
		"_ic_", "_il_", "_ir_",
		"_sc_", "_sl_", "_sr_",
		"_nc_", "_nl_", "_nr_",
		"_dc_", "_dl_", "_dr_",
	};

	private readonly string _connectionString;

	public DataProcessing(IConfiguration config)
	{
		_connectionString = config.GetConnectionString("DefaultConnection") ?? throw new Exception("Invalid Configuration");
	}

	private DataSet GetDataFromProcedure(string procedureName, Dictionary<string, string> procedureParamValue)
	{
		var ds = new DataSet();


		using var con = new SqlConnection(_connectionString);
		using var cmd = new SqlCommand(procedureName, con);
		cmd.CommandType = CommandType.StoredProcedure;

		var  ProcParams = GetProcedureParams(procedureName);
		foreach (ProcedureParam procParam in ProcParams)
		{
			

				string val = procedureParamValue.ContainsKey(procParam.Name) ? procedureParamValue[procParam.Name] : "NA";
				ProcessParameter(cmd, procParam, val);
		}

		if (param)
		{
			cmd.Parameters.Add("@user_id", SqlDbType.VarChar, 50).Value = "murthy";
		}
		using var da = new SqlDataAdapter(cmd);
		da.Fill(ds);

		return ds;
	}

	public List<ProcedureParam> GetProcedureParams(string procedureName)
	{
		var parameters = new List<ProcedureParam>();
		string query = @"
			SELECT 
				p.name AS ParamName, 
				t.name AS DataType,
				p.has_default_value,
				p.default_value
			FROM sys.parameters p
			INNER JOIN sys.types t ON p.user_type_id = t.user_type_id
			WHERE p.object_id = OBJECT_ID(@procedureName)";

		using (var con = new SqlConnection(_connectionString))
		using (var cmd = new SqlCommand(query, con))
		{
			cmd.Parameters.AddWithValue("@procedureName", procedureName);
			con.Open();
			using (var reader = cmd.ExecuteReader())
			{
				while (reader.Read())
				{
					parameters.Add(new ProcedureParam
					{
						Name = reader["ParamName"].ToString(),
						DataType = reader["DataType"].ToString(),
						HasDefaultValue = Convert.ToBoolean(reader["has_default_value"]),
						DefaultValue = reader["default_value"] != DBNull.Value ? reader["default_value"] : null
					});
				}
			}
		}
		return parameters;
	}

	private List<string> GetColumnKeysFromDataTable(DataTable dt)
	{
		return dt.Columns
				 .Cast<DataColumn>()
				 .Select(c => c.ColumnName)
				 .ToList();
	}

	private List<ColumnMeta> ParseColumns(List<string> keys)
	{
		var columns = new List<ColumnMeta>();

		foreach (var k in keys)
		{
			int typeIndex = -1;
			string datatypeMarker = null;
			for (int i = k.Length - 4; i >= 0; i--)
			{
				string sub = k.Substring(i, 4).ToLower();
				if (ValidMarkers.Contains(sub))
				{
					datatypeMarker = sub;
					typeIndex = i;
					break;
				}
			}

			if (datatypeMarker == null)
				throw new Exception($"Invalid column key: {k}. Must contain a marker like _ic_, _nr_, _sl_, etc.");

			string colNameRaw = k.Substring(0, typeIndex);
			string colName = colNameRaw.Replace("_", " ").ToUpper();

			char dataType = char.ToUpper(datatypeMarker[1]);
			char align = char.ToUpper(datatypeMarker[2]);

			int headerStart = typeIndex + 4;
			string header = headerStart < k.Length
				? k.Substring(headerStart).Replace("_", " ").ToUpper()
				: "DEFAULT";

			columns.Add(new ColumnMeta
			{
				ColumnKey = k,
				DisplayName = colName,
				DataType = dataType,
				Alignment = align,
				Header = header
			});
		}

		return columns;
	}

	private void GenerateExcelGroupHeader(List<ColumnMeta> columns, int currentCol, int headerRow, IXLWorksheet ws, int colHeaderRow)
	{
		foreach (var grp in columns.GroupBy(c => c.Header))
		{
			int startCol = currentCol;
			int count = grp.Count();

			ws.Range(headerRow, startCol, headerRow, startCol + count - 1).Merge();
			var headerCell = ws.Cell(headerRow, startCol);
			headerCell.Value = grp.Key;
			headerCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
			headerCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
			ws.Row(headerRow).Height = 30;
			ws.Cell(headerRow, startCol).Style.Font.SetBold();

			int i = 0;
			foreach (var col in grp)
			{
				int excelCol = startCol + i;

				col.ExcelColIndex = excelCol;

				var cell = ws.Cell(colHeaderRow, excelCol);
				cell.Value = col.DisplayName;
				cell.Style.Font.Bold = true;
				cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
				i++;
			}
			currentCol += count;
		}
		ws.Row(colHeaderRow).Height = 25;
	}

	private void GenerateExcelData(DataTable dt, List<ColumnMeta> columns, IXLWorksheet ws, int dataStartRow)
	{
		for (int row = 0; row < dt.Rows.Count; row++)
		{
			foreach (var meta in columns)
			{
				var cell = ws.Cell(dataStartRow + row, meta.ExcelColIndex);

				object value = dt.Rows[row][meta.ColumnKey];

				if (value == DBNull.Value)
				{
					cell.Value = "";
					continue;
				}

				switch (meta.DataType)
				{
					case 'I':
						cell.Value = Convert.ToInt32(value);
						cell.Style.NumberFormat.Format = "#,##0";
						break;

					case 'N':
						cell.Value = Convert.ToDouble(value);
						cell.Style.NumberFormat.Format = "#,##0.00";
						break;

					case 'S':
						cell.Value = value.ToString();
						break;

					case 'D':
						cell.Value = Convert.ToDateTime(value);
						cell.Style.DateFormat.Format = "yyyy-MM-dd";
						break;
				}

				cell.Style.Alignment.Horizontal = meta.Alignment switch
				{
					'L' => XLAlignmentHorizontalValues.Left,
					'C' => XLAlignmentHorizontalValues.Center,
					'R' => XLAlignmentHorizontalValues.Right,
					_ => XLAlignmentHorizontalValues.Left
				};
			}

		}
	}

	// Returns the filename for the generated sheet
	private string GenerateSheetName(string name)
	{
		return $"{name}.xlsx";
	}

	private void ProcessParameter(SqlCommand cmd, ProcedureParam procParam, string val)
	{
		switch (procParam.DataType.ToLower())
		{
			case "varchar":
			case "nvarchar":
			case "char":
			case "nchar":
			case "text":
			case "ntext":
				cmd.Parameters.Add(procParam.Name, SqlDbType.VarChar).Value = val ?? "NA";
				break;
			case "int":
			case "integer":
				if (int.TryParse(val, out int intVal))
					cmd.Parameters.Add(procParam.Name, SqlDbType.Int).Value = intVal;
				else
					cmd.Parameters.Add(procParam.Name, SqlDbType.Int).Value = DBNull.Value;
				break;
			case "bigint":
				if (long.TryParse(val, out long longVal))
					cmd.Parameters.Add(procParam.Name, SqlDbType.BigInt).Value = longVal;
				else
					cmd.Parameters.Add(procParam.Name, SqlDbType.BigInt).Value = DBNull.Value;
				break;
			case "decimal":
			case "numeric":
			case "money":
				if (decimal.TryParse(val, out decimal decVal))
					cmd.Parameters.Add(procParam.Name, SqlDbType.Decimal).Value = decVal;
				else
					cmd.Parameters.Add(procParam.Name, SqlDbType.Decimal).Value = DBNull.Value;
				break;
			case "float":
			case "real":
				if (double.TryParse(val, out double dblVal))
					cmd.Parameters.Add(procParam.Name, SqlDbType.Float).Value = dblVal;
				else
					cmd.Parameters.Add(procParam.Name, SqlDbType.Float).Value = DBNull.Value;
				break;
			case "datetime":
			case "date":
				if (DateTime.TryParse(val, out DateTime dateVal))
					cmd.Parameters.Add(procParam.Name, SqlDbType.DateTime).Value = dateVal;
				else
					cmd.Parameters.Add(procParam.Name, SqlDbType.DateTime).Value = DBNull.Value;
				break;
			case "bit":
			case "boolean":
				if (bool.TryParse(val, out bool boolVal))
					cmd.Parameters.Add(procParam.Name, SqlDbType.Bit).Value = boolVal;
				else if (val == "1")
					cmd.Parameters.Add(procParam.Name, SqlDbType.Bit).Value = true;
				else if (val == "0")
					cmd.Parameters.Add(procParam.Name, SqlDbType.Bit).Value = false;
				else
					cmd.Parameters.Add(procParam.Name, SqlDbType.Bit).Value = DBNull.Value;
				break;
			default:
				break;
		}
	}

	private string extractSheetNameFromProcedure(string name)
	{
		string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
		int lastDot = name.LastIndexOf('.');
		string procedureName = lastDot >= 0 ? name.Substring(lastDot + 1) : name;
		return $"{procedureName}_{timestamp}";
	}

	public (MemoryStream Stream, string FileName) GenerateExcelTemplate(string procedureName,  Dictionary<string, string> procedurParamValue)
	{
		var ds = GetDataFromProcedure(procedureName, procedurParamValue);

		if (ds.Tables.Count == 0)
			return (null, null);

		string filename = extractSheetNameFromProcedure(procedureName);
		string fullFileName = GenerateSheetName(filename);

		using var workbook = new XLWorkbook();

		bool hasData = false;
		for (int i = 0; i < ds.Tables.Count; i++)
		{
			var dt = ds.Tables[i];

			if (dt.Rows.Count == 0) continue;
			hasData = true;
			string sheetName = $"Sheet{i + 1}";
			var ws = workbook.Worksheets.Add(sheetName);

			var columnKeys = GetColumnKeysFromDataTable(dt);
			var columns = ParseColumns(columnKeys);

			int headerRow = 1;
			int colHeaderRow = 2;
			int dataStartRow = 3;
			GenerateExcelGroupHeader(columns, 1, headerRow, ws, colHeaderRow);
			GenerateExcelData(dt, columns, ws, dataStartRow);
			ws.Range(headerRow, 1, dataStartRow + dt.Rows.Count - 1, columns.Count)
			  .Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin)
			  .Border.SetInsideBorder(XLBorderStyleValues.Thin);
			ws.Columns().AdjustToContents();
			ws.SheetView.FreezeRows(2);
		}

		if (!hasData) return (null, null);

		var stream = new MemoryStream();
		workbook.SaveAs(stream);
		stream.Position = 0;
		return (stream, fullFileName);
	}

	public List<string> SearchProcedures(string query)
	{
		var procedures = new List<string>();
		string sql = @"
            SELECT TOP 7 
                s.name + '.' + p.name AS ProcedureName
            FROM sys.procedures p
            INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
            WHERE (s.name + '.' + p.name) LIKE @query
            ORDER BY p.name";

		using (var con = new SqlConnection(_connectionString))
		using (var cmd = new SqlCommand(sql, con))
		{
			cmd.Parameters.AddWithValue("@query", "%" + query + "%");
			con.Open();
			using (var reader = cmd.ExecuteReader())
			{
				while (reader.Read())
				{
					procedures.Add(reader["ProcedureName"].ToString());
				}
			}
		}
		return procedures;
	}

}
