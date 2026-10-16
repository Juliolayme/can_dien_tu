using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;

namespace Chuongtrinhcan;

internal class DataFetcher
{
	private static readonly DataFetcher instance = new DataFetcher();

	private readonly string connectionString;

	private OleDbCommand command;

	private OleDbCommand tranCommand;

	private DataTable table = null;

	private DataSet dataSet = null;

	private ArrayList mylist = null;

	private int nConString = 1;

	private OleDbConnection con = new OleDbConnection();

	public static DataFetcher Instance => instance;

	public DataFetcher()
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		connectionString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=sys/data.mdb;Jet OLEDB:Database Password=!@#^%$Nguyen1Van2A3;";
		((DbConnection)(object)con).ConnectionString = connectionString;
	}

	public string getConString()
	{
		return connectionString;
	}

	public string getDBPath()
	{
		int num = connectionString.LastIndexOf("=") + 1;
		return connectionString.Substring(num, connectionString.Length - num);
	}

	public void setConString(int nConString)
	{
		this.nConString = nConString;
	}

	public void setCommand(string query, int type)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		command = new OleDbCommand();
		((DbCommand)(object)command).CommandType = ((type == 1) ? CommandType.Text : CommandType.StoredProcedure);
		((DbCommand)(object)command).CommandText = query;
	}

	public void setTranCommand(string query, int type)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		tranCommand = new OleDbCommand();
		((DbCommand)(object)tranCommand).CommandType = ((type == 1) ? CommandType.Text : CommandType.StoredProcedure);
		((DbCommand)(object)tranCommand).CommandText = query;
	}

	public void addParameter(OleDbParameter aParam)
	{
		command.Parameters.Add(aParam);
	}

	public void addParameterWithValue(string param, string value)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		command.Parameters.Add(new OleDbParameter(param, (object)value));
	}

	public void addTranParameterWithValue(string param, string value)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		tranCommand.Parameters.Add(new OleDbParameter(param, (object)value));
	}

	public DataTable fetch()
	{
		//IL_0069: Expected O, but got Unknown
		command.Connection = con;
		try
		{
			if (((DbConnection)(object)con).State == ConnectionState.Closed)
			{
				((DbConnection)(object)con).Open();
			}
			table = new DataTable();
			table.Load((IDataReader)command.ExecuteReader(CommandBehavior.CloseConnection));
			((Component)(object)command).Dispose();
		}
		catch (OleDbException val)
		{
			OleDbException val2 = val;
			Console.WriteLine("Fetch Error: " + val2);
		}
		finally
		{
			if (((DbConnection)(object)con).State == ConnectionState.Open)
			{
				((DbConnection)(object)con).Close();
			}
		}
		return table;
	}

	public DataTable fetchTransaction(OleDbConnection conn, OleDbTransaction tran)
	{
		tranCommand.Connection = conn;
		tranCommand.Transaction = tran;
		if (((DbConnection)(object)conn).State == ConnectionState.Closed)
		{
			((DbConnection)(object)conn).Open();
		}
		table = new DataTable();
		table.Load((IDataReader)tranCommand.ExecuteReader(CommandBehavior.CloseConnection));
		((Component)(object)tranCommand).Dispose();
		return table;
	}

	public DataSet fetchDataSet()
	{
		//IL_0069: Expected O, but got Unknown
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_004c: Expected O, but got Unknown
		command.Connection = con;
		try
		{
			if (((DbConnection)(object)con).State == ConnectionState.Closed)
			{
				((DbConnection)(object)con).Open();
			}
			dataSet = new DataSet();
			OleDbDataAdapter val = new OleDbDataAdapter(command);
			((DataAdapter)(object)val).Fill(dataSet);
			((Component)(object)command).Dispose();
		}
		catch (OleDbException val2)
		{
			OleDbException val3 = val2;
			Console.WriteLine("FetchData Error: " + val3);
		}
		return dataSet;
	}

	public DataSet fetchDataSetWithTableName(string tableName)
	{
		//IL_006a: Expected O, but got Unknown
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_004c: Expected O, but got Unknown
		command.Connection = con;
		try
		{
			if (((DbConnection)(object)con).State == ConnectionState.Closed)
			{
				((DbConnection)(object)con).Open();
			}
			dataSet = new DataSet();
			OleDbDataAdapter val = new OleDbDataAdapter(command);
			((DbDataAdapter)(object)val).Fill(dataSet, tableName);
			((Component)(object)command).Dispose();
		}
		catch (OleDbException val2)
		{
			OleDbException val3 = val2;
			Console.WriteLine("FetchData Error: " + val3);
		}
		return dataSet;
	}

	public int execute()
	{
		//IL_0053: Expected O, but got Unknown
		command.Connection = con;
		int result = -1;
		try
		{
			if (((DbConnection)(object)con).State == ConnectionState.Closed)
			{
				((DbConnection)(object)con).Open();
			}
			result = ((DbCommand)(object)command).ExecuteNonQuery();
			((Component)(object)command).Dispose();
		}
		catch (OleDbException val)
		{
			OleDbException val2 = val;
			Console.WriteLine("Execute Error: " + val2);
		}
		return result;
	}

	public void executeTransaction(OleDbConnection conn, OleDbTransaction tran)
	{
		tranCommand.Connection = conn;
		tranCommand.Transaction = tran;
		if (((DbConnection)(object)conn).State == ConnectionState.Closed)
		{
			((DbConnection)(object)conn).Open();
		}
		((DbCommand)(object)tranCommand).ExecuteNonQuery();
		((Component)(object)tranCommand).Dispose();
	}

	public int getId()
	{
		//IL_0058: Expected O, but got Unknown
		command.Connection = con;
		int result = -1;
		try
		{
			if (((DbConnection)(object)con).State == ConnectionState.Closed)
			{
				((DbConnection)(object)con).Open();
			}
			result = Convert.ToInt32(((DbCommand)(object)command).ExecuteScalar());
			((Component)(object)command).Dispose();
		}
		catch (OleDbException val)
		{
			OleDbException val2 = val;
			Console.WriteLine("GetId Error: " + val2);
		}
		finally
		{
			if (((DbConnection)(object)con).State == ConnectionState.Open)
			{
				((DbConnection)(object)con).Close();
			}
		}
		return result;
	}

	public ArrayList getOneRow()
	{
		//IL_00a4: Expected O, but got Unknown
		command.Connection = con;
		try
		{
			if (((DbConnection)(object)con).State == ConnectionState.Closed)
			{
				((DbConnection)(object)con).Open();
			}
			OleDbDataReader val = command.ExecuteReader(CommandBehavior.SingleRow);
			object[] array = new object[((DbDataReader)(object)val).FieldCount];
			mylist = new ArrayList();
			if (((DbDataReader)(object)val).FieldCount > 0)
			{
				while (((DbDataReader)(object)val).Read())
				{
					((DbDataReader)(object)val).GetValues(array);
				}
			}
			mylist.AddRange(array);
			((Component)(object)command).Dispose();
		}
		catch (OleDbException val2)
		{
			OleDbException val3 = val2;
			Console.WriteLine("GetRow Error: " + val3);
		}
		finally
		{
			if (((DbConnection)(object)con).State == ConnectionState.Open)
			{
				((DbConnection)(object)con).Close();
			}
		}
		return mylist;
	}

	public string getString()
	{
		//IL_005c: Expected O, but got Unknown
		command.Connection = con;
		string result = "";
		try
		{
			if (((DbConnection)(object)con).State == ConnectionState.Closed)
			{
				((DbConnection)(object)con).Open();
			}
			result = ((DbCommand)(object)command).ExecuteScalar().ToString();
			((Component)(object)command).Dispose();
		}
		catch (OleDbException val)
		{
			OleDbException val2 = val;
			Console.WriteLine("GetId Error: " + val2);
		}
		finally
		{
			if (((DbConnection)(object)con).State == ConnectionState.Open)
			{
				((DbConnection)(object)con).Close();
			}
		}
		return result;
	}
}
