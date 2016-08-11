using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace SqliteToExcel
{
	class Program
	{
		static void Main(string[] args)
		{
			SQLiteConnection conn = new SQLiteConnection("Data Source = weather.db"); ;
			conn.Open();
			SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter("select * from nasaweather", conn);
			DataSet dataSet = new DataSet();
			dataAdapter.Fill(dataSet, "nasaweather");
			System.Data.DataTable dataTable = dataSet.Tables["nasaweather"];
			string path = Environment.CurrentDirectory + "\\";
			Application excel = new Application();
			Workbooks wbks = excel.Workbooks;
			Workbook wb = wbks.Add(path + "mb.xlsx");
			Worksheet wsh = wb.Sheets[1];
			int row = 2;
			foreach (DataRow dataRow in dataTable.Rows)
			{
				if (getName(dataRow["type"].ToString()) == "不要")
					continue;
				int col = 1;
				wsh.Cells[row, col++] = dataRow["lon"];
				wsh.Cells[row, col++] = dataRow["lat"];
				wsh.Cells[row, col++] = getName(dataRow["type"].ToString());
				string[] data = getArray(dataRow["data"].ToString());
				foreach (string str in data)
				{
					wsh.Cells[row, col++] = str;
				}
				//Console.WriteLine(dataRow["lat"] + " " + dataRow["lon"] + " " + dataRow["type"] + " " + dataRow["data"]);
				Console.WriteLine((row - 1.0)*100/13104 + "%");
				if ((row - 1) % 4 == 0)
				{
					Range lonRange = wsh.Range[wsh.Cells[row - 3, 1], wsh.Cells[row, 1]];
					lonRange.ClearContents();
					lonRange.Merge();
					wsh.Cells[row - 3, 1] = dataRow["lon"];
					Range latRange = wsh.Range[wsh.Cells[row - 3, 2], wsh.Cells[row, 2]];
					latRange.ClearContents();
					latRange.Merge();
					wsh.Cells[row - 3, 2] = dataRow["lat"];

					Range range = wsh.Range[wsh.Cells[row - 3, 1], wsh.Cells[row, col - 1]];
					range.BorderAround();

				}
				++row;
				//if (row > 100)
				//	break;
			}
			excel.DisplayAlerts = false;
			excel.AlertBeforeOverwriting = false;
			wb.SaveAs(path + "气象表.xlsx");
			wb.Close();
			wbks.Close();
		}

		static string getName(string input)
		{
			switch (input)
			{
				case "irradiance":
					return "辐照度";
				case "temperature":
					return "温度";
				case "humidity":
					return "湿度";
				case "wind":
					return "风速";
				default:
					return "不要";
			}
		}

		static string[] getArray(string input)
		{
			input = input.Replace("[", "");
			input = input.Replace("]", "");
			input = input.Replace(" ", "");
			string[] result = input.Split(',');
			if(result.Length != 12)
				throw new Exception("Split fail");
			for(int i = 0; i < result.Length; ++i)
				if (result[i] == "null")
					result[i] = "";
			return result;
		}
	}
}
