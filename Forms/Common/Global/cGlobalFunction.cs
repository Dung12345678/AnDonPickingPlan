using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BMS
{
	public static class cGlobalFunction
	{
		/// <summary>
		/// Trả về chuỗi query lấy toàn bộ dữ liệu của bảng
		/// </summary>
		/// <param name="tableName">Tên bảng</param>
		/// <param name="orderField">Tên trường sắp xếp</param>
		/// <returns></returns>
		public static string GetAllQuery(string tableName, string orderField) {
			string sql = string.Format("SELECT ROW_NUMBER() OVER (ORDER BY ID) AS STT, * FROM {0} ORDER BY {1}", tableName, orderField);
			return sql;
		}

		public static string GetWithSelectionQuery(string[] fieldName, string tableName, string orderField) {
			string listField = "";
			foreach (string item in fieldName) {
				listField = string.Concat(listField, ", ", item);
			}
			string sql = string.Format("SELECT ROW_NUMBER() OVER (ORDER BY ID) AS STT {0} FROM {1} ORDER BY {2}", listField, tableName, orderField);
			return sql;
		}

		
	}
}
