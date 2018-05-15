using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Data;
using System;
using System.Diagnostics;
using Microsoft.Win32;
using System.Linq;
using System.Threading;

namespace NOBEL.InstallHelper
{
    public static class SqlHelper
    {
        public static void InstallDatabase(string serverName)
        {            
            string scriptSql = GetDatabaseScript();
            IEnumerable<string> commandStrings = Regex.Split(scriptSql, @"^\s*GO\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);

            try
            {                
                DropDatabaseIfExist(serverName);
                CreateDatabase(serverName);
                ExecuteCommand(serverName, commandStrings);
                SaveServerNameToRegistry(serverName);
            }
            catch (Exception e)
            {
                throw new Exception("Have some errors while installing Nobel database. Please view the log file at C:\\InstallLog.txt", e);
            }
        }

        public static void UnInstallDatabase()
        {
            try
            {
                string serverName = GetServerNameFromRegistry();
                DropDatabaseIfExist(serverName);                
            }
            catch (Exception e)
            {
                throw new Exception("Have some errors while uninstallation Nobel database. Please view the log file at C:\\InstallLog.txt", e);
            }
        }

        public static bool DatabaseIsInstalled(string serverName = null)
        {
            try
            {  
                if(string.IsNullOrEmpty(serverName))
                {
                    serverName = GetServerNameFromRegistry();
                }

                RunCommand((commandFactory) =>
                {
                    var cmd = commandFactory();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select 1";
                    cmd.ExecuteScalar();
                }, serverName: serverName, databaseName: "NOBEL");

                return true;
            }
            catch(Exception)
            {
                return false;
            }
        }

        private static void CreateDatabase(string serverName)
        {
            const int timeOut = 20000;
            ProcessStartInfo pInfo = new ProcessStartInfo();
            pInfo.FileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\addselftosqlsysadmin.cmd";
            Process p = Process.Start(pInfo);
            p.WaitForExit(timeOut);

            if (p.HasExited)
            {
                Thread.Sleep(5000);
                RunCommand((commandFactory) =>
                {
                    var cmd = commandFactory();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = @"CREATE DATABASE [NOBEL]  COLLATE SQL_Latin1_General_CP1_CI_AS";
                    cmd.ExecuteNonQuery();
                }, serverName: serverName);
            }                        
        }

        private static void DropDatabaseIfExist(string serverName)
        {            
            RunCommand((commandFactory) =>
            {                
                var cmd = commandFactory();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'NOBEL')
                                    BEGIN           
                                        ALTER DATABASE [NOBEL] SET single_user with rollback immediate
                                        ALTER DATABASE [NOBEL] SET MULTI_USER                                   
                                        DROP DATABASE [NOBEL]
                                    END";
                cmd.ExecuteNonQuery();
            }, serverName: serverName);
        }

        private static void ExecuteCommand(string serverName, IEnumerable<string> scripts)
        {            
            RunCommand((commandFactory) =>
            {
                foreach (string script in scripts)
                {
                    if (!string.IsNullOrWhiteSpace(script))
                    {
                        var cmd = commandFactory();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = script;
                        cmd.ExecuteNonQuery();
                    }                    
                }                
            }, serverName: serverName, databaseName: "NOBEL");
        }
        private static void RunCommand(Action<Func<SqlCommand>> action, string serverName, string databaseName = "MASTER")
        {
            string connectionString = GetConnectionString(serverName, databaseName);
            SqlConnection.ClearAllPools();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();                
                action(() =>
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    return cmd;
                });                
            }
        }

        private static string GetConnectionString(string serverName, string databaseName)
        {
            if(string.IsNullOrEmpty(serverName))
            {
                serverName = "SQLEXPRESS";
            }

            if (serverName.ToUpper().Equals("SQLEXPRESS"))
            {
                serverName = $@".\{serverName}";
            }

            return $@"Server = {serverName}; Integrated security = SSPI; database = {databaseName}";                        
        }
        private static bool SaveServerNameToRegistry(string serverName)
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Nobel");
                key.SetValue("serverName", serverName);
                key.Close();

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static string GetServerNameFromRegistry()
        {
            try
            {
                var serverName = string.Empty;
                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Nobel");
                if (key != null)
                {
                    serverName = key.GetValue("serverName")?.ToString();
                    key.Close();                    
                }

                return serverName;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
        private static bool RemoveRegistry()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(@"SOFTWARE\Nobel");                
                return true;
            }
            catch(Exception)
            {
                return false;
            }
        }
        private static string GetDatabaseScript()
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(asm.GetName().Name + ".Nobel.sql"))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }                            
        }
    }
}
