using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
			DataTable dataTable = dataSet.Tables["nasaweather"];
			foreach (DataRow dataRow in dataTable.Rows)
			{
				Console.WriteLine(dataRow["lat"] + " " + dataRow["lon"] + " " + dataRow["type"] + " " + dataRow["data"]);
			}
		}
	}
}
