/*
Проверка наличия подключения в базе данных КМД 
и наличия учетной записи
*/

using System;
using System.Text;
using System.Data.SqlClient;
using System.Security.Principal;

namespace DBAccessChecker
{
    class CAccessChecker
    {
        private static string ConvertByteToStringSid(Byte[] sidBytes)
        {
            StringBuilder strSid = new StringBuilder();
            strSid.Append("S-");

            strSid.Append(sidBytes[0].ToString());
            if (sidBytes[6] != 0 || sidBytes[5] != 0)
            {
                string strAuth = String.Format
                ("0x{0:2x}{1:2x}{2:2x}{3:2x}{4:2x}{5:2x}",
                (Int16)sidBytes[1],
                (Int16)sidBytes[2],
                (Int16)sidBytes[3],
                (Int16)sidBytes[4],
                (Int16)sidBytes[5],
                (Int16)sidBytes[6]);
                strSid.Append("-");
                strSid.Append(strAuth);
            }
            else
            {
                Int64 iVal = (Int32)(sidBytes[1]) +
                (Int32)(sidBytes[2] << 8) +
                (Int32)(sidBytes[3] << 16) +
                (Int32)(sidBytes[4] << 24);
                strSid.Append("-");
                strSid.Append(iVal.ToString());
            }

            int iSubCount = Convert.ToInt32(sidBytes[7]);
            int idxAuth = 0;
            for (int i = 0; i < iSubCount; i++)
            {
                idxAuth = 8 + i * 4;

                if (idxAuth >= sidBytes.Length)
                    break;

                UInt32 iSubAuth = BitConverter.ToUInt32(sidBytes, idxAuth);
                strSid.Append("-");
                strSid.Append(iSubAuth.ToString());
            }

            return strSid.ToString();
        }

        private static SqlConnection GetDBConnection()
        {
            string _sconnect =
                "Data Source = apptestsrv;" +
                "Network Library = DBMSSOCN;" +
                "Initial Catalog = KMD;" +
                "Integrated Security = True";

            SqlConnection connection = new SqlConnection(_sconnect);

            return connection;
        }

        private static string GetDBUserSID(string name)
        {
            if (name == null) return (null);
            string _squery = "SELECT SID FROM rb_users WHERE UserLogin ='" + name + "'";
            object _reader = null;

            SqlConnection _connection = GetDBConnection();
            SqlCommand _command = new SqlCommand(_squery, _connection);

            try
            {
                _connection.Open();
                _reader = _command.ExecuteScalar();
                _connection.Close();
            }
            catch
            {
                return (null);
            }

            if (_reader == null) return (null);

            try
            {
                string _result = ConvertByteToStringSid((byte[])_reader);
                return (_result);
            }
            catch
            {
                return (null);
            }
        }

        private static string GetLocalSID()
        {
            try
            {
                var _windowsIdentity = WindowsIdentity.GetCurrent();
                return (_windowsIdentity.User.Value);
            }
            catch
            {
                return (null);
            }
        }

        private static string GetLocalName()
        {
            try
            {
                var _windowsIdentity = WindowsIdentity.GetCurrent();
                return (_windowsIdentity.Name);
            }
            catch
            {
                return (null);
            }
        }

        public static bool AccessCheck()
        {
            string _localName = GetLocalName();
            string _localSID = GetLocalSID();
            string _dbSID = GetDBUserSID(_localName);

            if (_localSID == null) return (false);
            if (_dbSID == null) return (false);
            if (_localSID != _dbSID) return (false);

            return (true);
        }
    }
}
