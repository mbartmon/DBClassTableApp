using System.Data;
using System.Windows.Forms;
using MbcsCentral;

public class clsDB
{
    LogClass Log = LogClass.instance;

	protected const string DS_DATA_SET = "wintrm";
	protected string loadStmt;
	protected string loadAllStmt;
	protected string insertStmt;
	protected string updateStmt;
	protected string localTable;
	protected string priKeyName;
	protected System.Data.SqlClient.SqlConnection localCN;

    protected string connectionString;
	
	protected string SQL_DELETE; // VBConversions Note: Initial value of ""DELETE FROM " + localTable + " WHERE " + priKeyName + " = ?;"" cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.

    public clsDB() { }

	public clsDB(string table, string priKeyName, string loadStmt, string loadAllStmt, string insertStmt, string updateStmt, System.Data.SqlClient.SqlConnection cn)
	{
		// VBConversions Note: Non-static class variable initialization is below.  Class variables cannot be initially assigned non-static values in C#.
		SQL_DELETE = "DELETE FROM " + localTable + " WHERE " + priKeyName + " = ?;";
		
		
		try
		{
			this.localTable = table;
			this.priKeyName = priKeyName;
			this.loadStmt = loadStmt;
			this.loadAllStmt = loadAllStmt;
			this.insertStmt = insertStmt;
			this.updateStmt = updateStmt;
			//this.localCN = cn;
            this.connectionString = cn.ConnectionString;
			SQL_DELETE = "DELETE FROM " + localTable + " WHERE [" + priKeyName + "] = ?;";
		}
		catch (System.Exception)
		{
			return;
		}
	}
	
	protected static string prepareStatement(string sql, string[] options)
	{
		string s = "";
		int i;
		int n = 0;
		
		if (options == null || (options.Length - 1) < 0)
		{
			s = sql;
		}
		else
		{
			for (i = 1; i <= sql.Length; i++)
			{
				if (sql.Substring(i - 1, 1) == "?")
				{
					if (n <= (options.Length - 1))
					{
						if (! (options[n] == null))
						{
							if (options[n].ToLower().Equals("true"))
							{
								options[n] = "1";
							}
							else if (options[n].ToLower().Equals("false"))
							{
								options[n] = "0";
							}
						}
						s += options[n];
						n++;
					}
					else
					{
						
					}
				}
				else
				{
					s += sql.Substring(i - 1, 1);
				}
			}
		}
		return s;
		
	}
	
	protected DataSet loadByPrimaryKey(string key)
	{
		
		string[] options = new string[2];
		options[0] = key;
		try
		{
			//localCN.Open();
            //localCN.ConnectionString = connectionString;
			string stmt = prepareStatement(loadStmt, options);
			System.Data.SqlClient.SqlDataAdapter da = new System.Data.SqlClient.SqlDataAdapter(stmt, this.connectionString);
			DataSet ds = new DataSet(DS_DATA_SET);
			da.Fill(ds, DS_DATA_SET);
			
			return ds;
		}
		catch (System.Data.SqlClient.SqlException ex)
		{
			MessageBox.Show("Error while retrieving record from " + localTable + ": " + ex.Message);
			return null;
		}
		catch (System.Exception ex)
		{
			Log.Log((int)LogClass.logType.ErrorCondition, "Error while retrieveing record from " + localTable + ": " + ex.Message);
			return null;
		}
		finally
		{
			try
			{
				localCN.Close();
			}
			catch
			{
			}
		}
		
	}
	
	protected DataSet loadAll()
	{
		
		/* if (localCN.State != ConnectionState.Open)
		{
			try
			{
				localCN.Open();
			}
			catch (System.Data.SqlClient.SqlException ex)
			{
				Log.Log((int)LogClass.logType.ErrorCondition, "Opening database: <" + localCN.ConnectionString + ">" + "\r\n" + "   " + loadAllStmt + "\r\n" + "   " + ex.Message, false, MessageBoxButtons.OK);
				return null;
			}
			catch (System.Exception ex)
			{
				Log.Log((int)LogClass.logType.ErrorCondition, "Opening database: <" + localCN.ConnectionString + ">" + "\r\n" + "   " + loadAllStmt + "\r\n" + "   " + ex.Message);
				return null;
			}
		}   */
		try
		{
            //localCN.ConnectionString = connectionString;
            System.Data.SqlClient.SqlDataAdapter da = new System.Data.SqlClient.SqlDataAdapter(loadAllStmt, this.connectionString);
			DataSet ds = new DataSet(DS_DATA_SET);
			da.Fill(ds, DS_DATA_SET);
			return ds;
		}
		catch (System.Data.SqlClient.SqlException ex)
		{
			Log.Log((int)LogClass.logType.ErrorCondition, "Error while retrieving records from " + localTable + ": " + ex.Message, false, MessageBoxButtons.OK);
			return null;
		}
		catch (System.Exception ex)
		{
			Log.Log((int)LogClass.logType.ErrorCondition, "Error while retrieving records from " + localTable + ": " + ex.Message, false, MessageBoxButtons.OK);
		}
		finally
		{
			try
			{
				
				localCN.Close();
			}
			catch
			{
			}
		}
		return null;
		
	}
	
	protected DataSet load(string stmt, string[] options)
	{
		
		if (options != null)
		{
			if (options.Length > 0 && options[0].Trim() != "")
			{
				stmt = prepareStatement(stmt, options);
			}
		}
        /*
		if (localCN.State != ConnectionState.Open)
		{
			try
			{
				localCN.Open();
			}
			catch (System.Exception)
			{
				return null;
			}
		}
         */
        System.Data.SqlClient.SqlDataAdapter da = new System.Data.SqlClient.SqlDataAdapter(stmt, this.connectionString);
		DataSet ds = new DataSet(DS_DATA_SET);
		try
		{
			da.Fill(ds, DS_DATA_SET);
            /*try
            {
                localCN.Close();
            }
            catch
            {
            } */
        }
		catch (System.Data.SqlClient.SqlException)
		{
			return null;
		} 
		return ds;
		
	}
	
	protected bool updateRecord(string key, string[] values)
	{
		try
        {
            try
            {
                localCN = new System.Data.SqlClient.SqlConnection(connectionString);
                localCN.Open();
            }
            catch { }  
            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
			cmd.CommandType = CommandType.Text;
			cmd.Connection = localCN;
			cmd.CommandText = prepareStatement(this.updateStmt, values);
			cmd.ExecuteNonQuery();
		}
		catch (System.Data.SqlClient.SqlException ex)
		{
			MessageBox.Show("Error while updating record: " + ex.Message);
			return false;
		}
		finally
		{
			try
			{
				localCN.Close();
			}
			catch
			{
			}
		}
		return true;
	}
	
	protected int insertRecord(string[] values)
	{
		int id = -1;
		try
		{
            try
            {
                localCN = new System.Data.SqlClient.SqlConnection( connectionString);
                localCN.Open();
            }
            catch (System.Exception ex)
            {
            	
            }
            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
			cmd.CommandType = CommandType.Text;
			cmd.Connection = localCN;
			cmd.CommandText = prepareStatement(this.insertStmt, values);
			cmd.ExecuteNonQuery();
			cmd.CommandText = "SELECT @@Identity";
			id = System.Convert.ToInt32(cmd.ExecuteScalar());
			//id = 0
		}
        catch (System.Data.SqlClient.SqlException ex)
		{
			if (ex.Message.ToLower().IndexOf("duplicate") + 1 > 0)
			{
				MessageBox.Show("Attempt to add a duplicate record.");
			}
			else
			{
				MessageBox.Show("Error while inserting record into " + localTable + ": " + ex.Message);
			}
			return id;
		}
		finally
		{
			try
			{
				localCN.Close();
			}
			catch
			{
			}
		}
		return id;
	}
	
	public bool delete(string key)
	{
		
		try
		{
            localCN = new System.Data.SqlClient.SqlConnection( connectionString);
            localCN.Open();
			string[] values = new string[] {key};
            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
			cmd.CommandType = CommandType.Text;
			cmd.Connection = localCN;
			cmd.CommandText = prepareStatement(SQL_DELETE, values);
			cmd.ExecuteNonQuery();
		}
        catch (System.Data.SqlClient.SqlException ex)
		{
			MessageBox.Show("Error while deleting record: " + ex.Message);
			return false;
		}
		finally
		{
			try
			{
				localCN.Close();
			}
			catch
			{
			}
		}
		return true;
		
	}

    protected bool query(string stmt, string[] options)
    {

        try
        {
            localCN = new System.Data.SqlClient.SqlConnection( connectionString);
            localCN.Open();
        }
        catch (System.Exception ex)
        {
            GlobalShared.Log.Log((int)LogClass.logType.ErrorCondition, "Unable to open DB connection <" + localCN.ConnectionString + "> for query <" + stmt + ">: " + ex.Message);
            return false;
        }
        try
        {
            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
            cmd.CommandType = CommandType.Text;
            cmd.Connection = localCN;
            cmd.CommandText = prepareStatement(stmt, options);
            cmd.ExecuteNonQuery();
        }
        catch (System.Data.SqlClient.SqlException ex)
        {
            GlobalShared.Log.Log((int)LogClass.logType.ErrorCondition, "Executing Query <" + stmt + ">: " + ex.Message);
            return false;
        }
        finally
        {
            try
            {
                localCN.Close();
            }
            catch { }
        }
 
        return true;

    }

    public static DataSet Query(string stmt, System.Data.SqlClient.SqlConnection localCN)
    {

        System.Data.SqlClient.SqlDataAdapter da = new System.Data.SqlClient.SqlDataAdapter(stmt, localCN.ConnectionString);
        DataSet ds = new DataSet(DS_DATA_SET);
        try
        {
            da.Fill(ds, DS_DATA_SET);
            try
            {
                localCN.Close();
            }
            catch
            {
            }
        }
        catch (System.Data.SqlClient.SqlException)
        {
            return null;
        }
        return ds;

    }

    public System.Data.SqlClient.SqlConnection LocalCN
    {
        get { return localCN; }
        set { localCN = value;
        this.connectionString = LocalCN.ConnectionString;
        }
    }
}