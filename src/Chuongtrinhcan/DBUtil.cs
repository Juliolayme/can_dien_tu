using System;
using System.Data;
using System.Drawing;
using System.Management;
using System.Windows.Forms;

namespace Chuongtrinhcan;

internal class DBUtil
{
	private static DataFetcher dataFetcher = DataFetcher.Instance;

	public static bool isDuplicated(string value, string field, string table)
	{
		string query = "SELECT * FROM " + table + " WHERE " + field + " = '" + value + "'";
		dataFetcher.setCommand(query, 1);
		DataTable dataTable = dataFetcher.fetch();
		if (dataTable.Rows.Count > 0)
		{
			return true;
		}
		return false;
	}

	public static bool isDuplicatedUpdate(string value, string field, string table, string idField, string idValue)
	{
		string query = "SELECT * FROM " + table + " WHERE " + field + " = '" + value + "' and " + idField + "<>" + idValue;
		dataFetcher.setCommand(query, 1);
		DataTable dataTable = dataFetcher.fetch();
		if (dataTable.Rows.Count > 0)
		{
			return true;
		}
		return false;
	}

	public static string getValue(string sql, string param)
	{
		dataFetcher.setCommand(sql, 1);
		dataFetcher.addParameterWithValue("@param", param);
		return dataFetcher.getString();
	}

	public static string getValueNoParam(string sql)
	{
		dataFetcher.setCommand(sql, 1);
		return dataFetcher.getString();
	}

	public static void loadCombo(ComboBox cbo, string sql, string fieldDisplay)
	{
		dataFetcher.setCommand(sql, 1);
		DataSet dataSet = dataFetcher.fetchDataSet();
		cbo.DataSource = dataSet.Tables[0];
		((ListControl)cbo).DisplayMember = fieldDisplay;
		((ListControl)cbo).ValueMember = fieldDisplay;
	}

	public static void loadComboValue(ComboBox cbo, string sql, string fieldDisplay, string fieldValue)
	{
		dataFetcher.setCommand(sql, 1);
		DataSet dataSet = dataFetcher.fetchDataSet();
		cbo.DataSource = dataSet.Tables[0];
		((ListControl)cbo).DisplayMember = fieldDisplay;
		((ListControl)cbo).ValueMember = fieldValue;
	}

	public static void LoadMaunen(Form frm)
	{
		try
		{
			switch (getValueNoParam("select maunen from thongsohethong"))
			{
			case "Đỏ":
				((Control)frm).BackColor = Color.Red;
				break;
			case "Vàng":
				((Control)frm).BackColor = Color.Yellow;
				break;
			case "Xanh":
				((Control)frm).BackColor = Color.Green;
				break;
			case "Trắng":
				((Control)frm).BackColor = Color.White;
				break;
			case "Cam":
				((Control)frm).BackColor = Color.Orange;
				break;
			case "Hồng":
				((Control)frm).BackColor = Color.Pink;
				break;
			default:
				((Control)frm).BackColor = Color.FromArgb(0, 192, 192);
				break;
			}
		}
		catch
		{
		}
	}

	public static string getCPUID()
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Expected O, but got Unknown
		string text = "";
		ManagementClass val = new ManagementClass("win32_processor");
		ManagementObjectCollection instances = val.GetInstances();
		ManagementObjectEnumerator enumerator = instances.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ManagementObject val2 = (ManagementObject)enumerator.Current;
				if (text == "")
				{
					text = ((ManagementBaseObject)val2).Properties["processorID"].Value.ToString();
					break;
				}
			}
		}
		finally
		{
			((IDisposable)enumerator)?.Dispose();
		}
		return text;
	}

	public static double ConvertToTotalNumber(string duLieu)
	{
		double num = 0.0;
		char[] array = duLieu.ToCharArray();
		for (int i = 0; i < array.Length; i++)
		{
			try
			{
				num += double.Parse(array[i].ToString());
			}
			catch
			{
			}
		}
		return num;
	}
}
