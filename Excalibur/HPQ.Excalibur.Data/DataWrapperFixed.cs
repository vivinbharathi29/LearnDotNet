using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace HPQ.Data
{
    //[Obsolete("this class is not supposed to be used in any new code, it has been moved here and heavily modified just to make the legacy excelexport run without leaking connections. As a result it's poorly coded and it has side effects for any other use.")]
    public class DataWrapper
    {

        public DataWrapper()
        {
            _connStr = ConfigurationManager.ConnectionStrings["PRSConnectionString"].ConnectionString;
        }

        private string _connStr = null;

        public int ExecuteCommandNonQuery(SqlCommand cmd)
        {
            using (var cn = new SqlConnection(_connStr))
            {
                cmd.Connection = cn;
                cn.Open();
                int x = cmd.ExecuteNonQuery();
                cmd.Dispose();
                return x;
            }
        }
        public int ExecuteSqlNonQuery(string cmdText)
        {
            using (var cn = new SqlConnection(_connStr))
            using (var cmd = cn.CreateCommand())
            {
                cmd.Connection = cn;
                cmd.CommandText = cmdText;
                cn.Open();
                int x = cmd.ExecuteNonQuery();
                cmd.Dispose();
                return x;
            }
        }

        public DataTable ExecuteCommandTable(SqlCommand cmd)
        {
            DataSet ds = new DataSet();

            using (var cn = new SqlConnection(_connStr))
            {
                cmd.Connection = cn;
                var da = new SqlDataAdapter(cmd);
                cn.Open(); //probably not necessary with dataadapters
                da.Fill(ds);
                da.Dispose();
                cmd.Dispose();
            }

            if (ds.Tables.Count > 0)
                return ds.Tables[0];
            else
                return new DataTable();
        }

        public SqlCommand CreateCommand(string procName)
        {
            return new SqlCommand(procName) { CommandType = CommandType.StoredProcedure };
        }

        public SqlCommand CreateCommand(string cmdText, System.Data.CommandType cmdType)
        {
            if (cmdType == CommandType.StoredProcedure)
                return new SqlCommand(cmdText) { CommandType = CommandType.StoredProcedure };

            return new SqlCommand(cmdText) { CommandType = CommandType.Text };
        }


        private void CreateParameter(SqlCommand cmd, string paramName, SqlDbType paramType, string paramValue, int paramSize, int paramScale, ParameterDirection paramDirection = ParameterDirection.Input)
        {
            SqlParameter param = new SqlParameter(paramName, paramType);
            param.Direction = paramDirection;

            //KB 2011/05/03 - Seperated the size, precision & scale values out so they could be set for output variables when the input value was null.
            switch (paramType)
            {
                case SqlDbType.Char:
                case SqlDbType.NChar:
                case SqlDbType.VarChar:
                case SqlDbType.NVarChar:
                case SqlDbType.Text:
                case SqlDbType.NText:
                    if (paramSize != 0) param.Size = paramSize;
                    break;
                case SqlDbType.Decimal:
                    param.Precision = Convert.ToByte(paramSize);
                    param.Scale = Convert.ToByte(paramScale);
                    break;
                default:
                    if (paramSize != 0) param.Size = paramSize;
                    break;
            }

            if (paramValue == String.Empty || paramValue == null)
            {
                // Substitute a null for an empty argument
                param.Value = DBNull.Value;
            }
            else
            {
                switch (paramType)
                {
                    case SqlDbType.Char:
                    case SqlDbType.NChar:
                    case SqlDbType.VarChar:
                    case SqlDbType.NVarChar:
                    case SqlDbType.Text:
                    case SqlDbType.NText:
                        param.Value = paramValue;
                        break;
                    case SqlDbType.BigInt:
                        try
                        {
                            param.Value =
                            Convert.ToInt64(Decimal.Parse(paramValue.Replace("$", ""),
                            NumberStyles.Any));
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value '" + paramValue + "' is not a number.", err);
                        }
                        break;
                    case SqlDbType.Bit:
                        try
                        {
                            switch (paramValue.ToUpper())
                            {
                                case "T":
                                case "TRUE":
                                case "Y":
                                case "YES":
                                case "1":
                                    param.Value = true;
                                    break;
                                case "F":
                                case "FALSE":
                                case "N":
                                case "NO":
                                case "0":
                                    param.Value = false;
                                    break;
                                default:
                                    param.Value = Boolean.Parse(paramValue);
                                    break;
                            }
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value '" + paramValue + "' is not a TRUE/FALSE value.", err);
                        }
                        break;
                    case SqlDbType.Int:
                        try
                        {
                            param.Value =
                            Convert.ToInt32(Decimal.Parse(paramValue.Replace("$", ""),
                            NumberStyles.Any));
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value '" + paramValue + "' is not a number.", err);
                        }
                        break;
                    case SqlDbType.SmallInt:
                        try
                        {
                            param.Value =
                            Convert.ToInt16(Decimal.Parse(paramValue.Replace("$", ""),
                            NumberStyles.Any));
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value '" + paramValue + "' is not a valid number.", err);
                        }
                        break;
                    case SqlDbType.TinyInt:
                        try
                        {
                            param.Value =
                            Convert.ToByte(Decimal.Parse(paramValue.Replace("$", ""),
                            NumberStyles.Any));
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value '" + paramValue + "' is not a valid number.", err);
                        }
                        break;
                    case SqlDbType.DateTime:
                        try
                        {
                            param.Value = DateTime.Parse(paramValue);
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value '" + paramValue + "' is not a valid date.", err);
                        }
                        break;

                    case SqlDbType.Decimal:
                        try
                        {
                            param.Value = Decimal.Parse(paramValue.Replace("$", ""),
                            NumberStyles.Any);
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value '" + paramValue + "' is not a number.", err);
                        }
                        break;
                    case SqlDbType.Float:
                    case SqlDbType.Money:
                    case SqlDbType.Real:
                    case SqlDbType.SmallMoney:
                        try
                        {
                            param.Value = Convert.ToDouble(Decimal.Parse(paramValue.Replace("$", ""), NumberStyles.Any));
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value '" + paramValue + "' is not a number.", err);
                        }
                        break;
                    case SqlDbType.Binary:
                    case SqlDbType.VarBinary:
                        try
                        {
                            param.Value = Convert.FromBase64String(paramValue);
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value must be a Base64-encoded string.", err);
                        }
                        break;
                    case SqlDbType.UniqueIdentifier:
                        try
                        {
                            param.Value = new Guid(paramValue);
                        }
                        catch (Exception err)
                        {
                            throw new Exception("Parameter value must be a valid GUID.", err);
                        }
                        break;
                    default:
                        param.Value = paramValue;
                        break;
                }
            }
            cmd.Parameters.Add(param);
        }

        public void CreateParameter(SqlCommand cmd, string paramName, SqlDbType paramType, string paramValue, int paramSize)
        {
            CreateParameter(cmd, paramName, paramType, paramValue, paramSize, 0);
        }

        public void CreateParameter(SqlCommand cmd, string paramName, SqlDbType paramType, string paramValue)
        {
            CreateParameter(cmd, paramName, paramType, paramValue, 0, 0);
        }

        public void CreateParameter(SqlCommand cmd, string paramName, SqlDbType paramType, string paramValue, ParameterDirection paramDirection)
        {
            CreateParameter(cmd, paramName, paramType, paramValue, 0, 0, paramDirection);
        }

        public void CreateParameter(SqlCommand cmd, string paramName, SqlDbType paramType, string paramValue, int paramSize, ParameterDirection paramDirection)
        {
            CreateParameter(cmd, paramName, paramType, paramValue, paramSize, 0, paramDirection);
        }

        public object ExecuteCommandScalar(string sql)
        {
            using (var cn = new SqlConnection(_connStr))
            using (var cmd = cn.CreateCommand())
            {
                cmd.CommandText = sql;
                cn.Open();
                return cmd.ExecuteScalar();
            }

        }

        public object ExecuteCommandScalarSP(string spName, params SqlParameter[] pars)
        {
            using (var cn = new SqlConnection(_connStr))
            using (var cmd = cn.CreateCommand())
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = spName;
                if (pars != null)
                    cmd.Parameters.AddRange(pars);
                cn.Open();
                return cmd.ExecuteScalar();
            }

        }

        public object ExecuteCommandScalar(SqlCommand cmd)
        {
            using (var cn = new SqlConnection(_connStr))
            {
                cmd.Connection = cn;
                cn.Open();
                return cmd.ExecuteScalar();
            }

        }
    }

}
