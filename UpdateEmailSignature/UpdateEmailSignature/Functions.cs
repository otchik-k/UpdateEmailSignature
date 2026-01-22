using Microsoft.Data.SqlClient;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Globalization;


public class Functions
{
    public static List<string> CutNullData(string[] anyArray)
    {
        List<string> cat = new List<string>();
        for (int i = 0; i < anyArray.Length; i++)
        {
            if (anyArray[i] != null)
            {
                cat.Add(anyArray[i].ToString());
            }
        }
        return cat;
    }


    public static Dictionary<string, string> GetAdUserAtributs(string ldapPath, string login)
    {
        Dictionary<string, string> resultDict = new Dictionary<string, string>
        {
            {"cn", "" },
            {"title", "" },
            {"company", "" },
            {"streetAddress", "" },
            {"l", "" },
            {"mail", "" },
            {"telephoneNumber", "" },
            {"sAMAccountName", "" },
            {"mobile", "" },
            {"pager", "" }
        };

        using (DirectorySearcher searcher = new DirectorySearcher())
        {
            searcher.SearchRoot = new DirectoryEntry("LDAP://" + ldapPath);
            searcher.Filter = $"sAMAccountName={login}";

            searcher.PropertiesToLoad.Add("cn");
            searcher.PropertiesToLoad.Add("title");
            searcher.PropertiesToLoad.Add("company");
            searcher.PropertiesToLoad.Add("streetAddress");
            searcher.PropertiesToLoad.Add("l");
            searcher.PropertiesToLoad.Add("mail");
            searcher.PropertiesToLoad.Add("telephoneNumber");
            searcher.PropertiesToLoad.Add("mobile");
            searcher.PropertiesToLoad.Add("pager");
            

            SearchResult result = searcher.FindOne();

            if (result != null)
            {
                foreach (var key in resultDict.Keys)
                {
                    if (result.Properties.Contains(key) && result.Properties[key].Count > 0)
                    {
                        resultDict[key] = result.Properties[key][0].ToString();
                    }
                    else
                    {
                        resultDict[key] = null;
                    }
                }
            }
            else
            {
                UsingStreamWriter("Пользователь" + login + " не найден.");
            }
        }
        return resultDict;
    }


    public static List<string> GetAllLoginAd (string ldapPath)
    {
        
        using (DirectoryEntry entry = new DirectoryEntry("LDAP://" + ldapPath))
        {
            try
            {
                DirectorySearcher searcher = new DirectorySearcher(entry);
                searcher.Filter = "(objectClass=user)";
                searcher.PropertiesToLoad.Add("sAMAccountName");
                SearchResultCollection results = searcher.FindAll();

                List<string> samAccountNames = new List<string>();
                foreach (SearchResult result in results)
                {
                    if (result.Properties.Contains("sAMAccountName"))
                    {
                        string samAccountName = (string)result.Properties["sAMAccountName"][0];
                        samAccountNames.Add(samAccountName);
                    }

                }

                return samAccountNames;
            }
            catch (Exception ex)
            {
                UsingStreamWriter (ex.Message);
                return null;
            }
            
        }
    }


    public static string GetSqlIdentityID(string connectToSQL, string cn, string email)
    {
        string result = "";
        using (SqlConnection connection = new SqlConnection(connectToSQL))
        {
            connection.Open();
            string queryText = $"SELECT identity_id FROM identities WHERE name=\'{cn}\' AND email=\'{email}\'";
            using (SqlCommand command = new SqlCommand(queryText, connection))
            {
                try
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result = reader.GetInt32(0).ToString();
                        }
                    }
                }
                catch (Exception ex)
                {
                    UsingStreamWriter(ex.Message);
                    result = null;
                }
                
            }
        }
        return result;
    }


    public static void UpdateSqLSignature (string connectToSQL, string signatureText, string identity_id)
    {
        using (SqlConnection con = new SqlConnection(connectToSQL))
        {
            con.Open();
            string queryText = "UPDATE identities SET signature=@signature, html_signature=1  WHERE identity_id=@identity_id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(queryText, con))
                {
                    cmd.Parameters.AddWithValue("@signature", signatureText);
                    cmd.Parameters.AddWithValue("@identity_id", identity_id);

                    int rowsAffected = cmd.ExecuteNonQuery();
                }
            }
            catch (SqlException ex)
            {
               UsingStreamWriter(ex.Message);
            }
        }
    }


    public static string GetSqlSignature(string connectToSQL, string identity_id)
    {
        string result = "";
        try
        {
            using (SqlConnection connection = new SqlConnection(connectToSQL))
            {
                connection.Open();
                string queryText = "SELECT signature FROM identities WHERE identity_id=@identity_id";
                using (SqlCommand command = new SqlCommand(queryText, connection))
                {
                    command.Parameters.AddWithValue("@identity_id", identity_id.ToString());
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result = reader.GetString(0);
                        }
                    }
                }
            }
            return result;
        }
        catch(Exception ex)
        {
            UsingStreamWriter (ex.Message);
            return null;
        }
    }


    public static List<string> GetSqlAllMail(string coccectToSQL)
    {
        List<string> result = new List<string>();
        try
        {
            using (SqlConnection connection = new SqlConnection(coccectToSQL))
            {
                connection.Open();
                string queryText = "SELECT email FROM identities";
                using (SqlCommand command = new SqlCommand (queryText, connection))
                {
                    using(SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                result.Add(reader.GetString(0));
                            }
                        }
                    }
                }
            }
        }
        catch(Exception ex)
        {
            UsingStreamWriter(ex.Message);
        }
        return result;
    }


    public static void FindeNameDoesntExistAd(string coccectToSQL, string mail, string cn)
    {
        List<string> nameFromSql = new List<string>();
        try
        {
            using (SqlConnection connection = new SqlConnection(coccectToSQL))
            {
                connection.Open();
                string queryText = $"SELECT name FROM identities WHERE email=\'{mail}\'";
                using (SqlCommand command = new SqlCommand (queryText , connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            nameFromSql.Add(reader.GetString(0));
                        }
                    }
                }
            }
            
            for (int i = 0; i < nameFromSql.Count; i++)
            {
                if (nameFromSql[i] != cn)
                {
                    UsingStreamWriter("mail=" + mail);
                    UsingStreamWriter("В AD cn=" + cn);
                    UsingStreamWriter("В SQL name=" + nameFromSql[i]);
                    UsingStreamWriter("");
                }
            }

        }
        catch (Exception ex)
        {
            UsingStreamWriter (ex.Message);
        }
    }


    public static string[] GetParametrValue(string configFileLine, string separator)
    {
        return configFileLine.Split(separator);
    }

    public static string logName = "log\\log - " + DateTime.Now.ToString("dd.MM.yyyy HH.mm.ss") + ".txt";
    public static void UsingStreamWriter(string anyString)
    {
        using (StreamWriter logWriter = new StreamWriter(logName, true))
        {
            logWriter.WriteLine(anyString);
        }
    }
}
    
