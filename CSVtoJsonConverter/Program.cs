using System;
using System.IO;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;

namespace CSVtoJsonConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // ReadTwocolumnCSV();
            // Read Excel file and convert it into Data Table
            DataTable table = ConvertExcelToDataTable("i18nEHupdate.xlsx");

            // Read the datatable and insert into List of object.
            List<Languages> languages = new List<Languages>();
            for (int index = 0; index < table.Rows.Count; index++)
            {
                Languages lan = new Languages();
                lan.Key = table.Rows[index]["Key"].ToString();
                lan.English = table.Rows[index]["English"].ToString();
                lan.Norwegian = table.Rows[index]["Norwegian"].ToString();
                lan.Danish = table.Rows[index]["Danish"].ToString();
                lan.Swedish = table.Rows[index]["Swedish"].ToString();
                languages.Add(lan);
            }

            ConvertToEnglish(languages);
            ConvertToNorwegian(languages);
            ConvertToSwedish(languages);
            ConvertToDanish(languages);
            Console.ReadLine();
        }

        private static void ConvertToDanish(List<Languages> languages)
        {
            List<string> uniqueFirst = new List<string>();
            Dictionary<string, string> uniqueSecond = new Dictionary<string, string>();
            Dictionary<string, string> properties = new Dictionary<string, string>();
            string en = $"{{ {Environment.NewLine}";
            foreach (var item in languages)
            {
                string[] keys = item.Key.Split('.');

                uniqueFirst.Add(keys[0]);
                if (keys.Length == 2 && !properties.ContainsKey(keys[1]))
                {
                    properties.Add(keys[1], keys[0]);
                }
                else if (keys.Length == 3 && !uniqueSecond.ContainsKey(keys[1]))
                {
                    uniqueSecond.Add(keys[1], keys[0]);
                    if (!properties.ContainsKey(keys[2]))
                        properties.Add(keys[2], keys[1]);
                }
            }
            uniqueFirst = uniqueFirst.Distinct<string>().ToList();

            foreach (string key in uniqueFirst)
            {
                en += $"\"{key}\" : {{{Environment.NewLine}";
                var firstChild = uniqueSecond.Where(a => a.Value == key).Select(s => s.Key).ToList();
                if (firstChild != null && firstChild.Count > 0)
                {
                    foreach (string child in firstChild)
                    {
                        en += $"\"{child}\" : {{{Environment.NewLine}";
                        var firstChildProperties = languages.Where(a => a.Key.Contains(key + "." + child + "."));
                        foreach (var item in firstChildProperties)
                        {
                            en += $"\"{item.Key.Split('.')[2]}\":\"{item.Danish}\",{Environment.NewLine}";
                        }
                        en += $"}},{Environment.NewLine}";
                    }

                }
                var children = properties.Where(a => a.Value == key).Select(s => s.Key).ToList();
                if (children != null && children.Count > 0)
                {
                    foreach (string child in children)
                    {
                        var item = languages.FirstOrDefault(a => a.Key.Equals(key + "." + child));
                        if (item != null)
                        {
                            en += $"\"{item.Key.Split('.')[1]}\":\"{item.Danish}\",{Environment.NewLine}";
                        }
                    }

                }
                en += $"}},{Environment.NewLine}";
            }

            en += "}";
            File.WriteAllText("Danish.json", en, Encoding.UTF8);
        }

        private static void ConvertToSwedish(List<Languages> languages)
        {
            List<string> uniqueFirst = new List<string>();
            Dictionary<string, string> uniqueSecond = new Dictionary<string, string>();
            Dictionary<string, string> properties = new Dictionary<string, string>();
            string en = $"{{ {Environment.NewLine}";
            foreach (var item in languages)
            {
                string[] keys = item.Key.Split('.');

                uniqueFirst.Add(keys[0]);
                if (keys.Length == 2 && !properties.ContainsKey(keys[1]))
                {
                    properties.Add(keys[1], keys[0]);
                }
                else if (keys.Length == 3 && !uniqueSecond.ContainsKey(keys[1]))
                {
                    uniqueSecond.Add(keys[1], keys[0]);
                    if (!properties.ContainsKey(keys[2]))
                        properties.Add(keys[2], keys[1]);
                }
            }
            uniqueFirst = uniqueFirst.Distinct<string>().ToList();

            foreach (string key in uniqueFirst)
            {
                en += $"\"{key}\" : {{{Environment.NewLine}";
                var firstChild = uniqueSecond.Where(a => a.Value == key).Select(s => s.Key).ToList();
                if (firstChild != null && firstChild.Count > 0)
                {
                    foreach (string child in firstChild)
                    {
                        en += $"\"{child}\" : {{{Environment.NewLine}";
                        var firstChildProperties = languages.Where(a => a.Key.Contains(key + "." + child + "."));
                        foreach (var item in firstChildProperties)
                        {
                            en += $"\"{item.Key.Split('.')[2]}\":\"{item.Swedish}\",{Environment.NewLine}";
                        }
                        en += $"}},{Environment.NewLine}";
                    }

                }
                var children = properties.Where(a => a.Value == key).Select(s => s.Key).ToList();
                if (children != null && children.Count > 0)
                {
                    foreach (string child in children)
                    {
                        var item = languages.FirstOrDefault(a => a.Key.Equals(key + "." + child));
                        if (item != null)
                        {
                            en += $"\"{item.Key.Split('.')[1]}\":\"{item.Swedish}\",{Environment.NewLine}";
                        }
                    }

                }
                en += $"}},{Environment.NewLine}";
            }

            en += "}";
            File.WriteAllText("Swedish.json", en, Encoding.UTF8);
        }

        private static void ConvertToNorwegian(List<Languages> languages)
        {
            List<string> uniqueFirst = new List<string>();
            Dictionary<string, string> uniqueSecond = new Dictionary<string, string>();
            Dictionary<string, string> properties = new Dictionary<string, string>();
            string en = $"{{ {Environment.NewLine}";
            foreach (var item in languages)
            {
                string[] keys = item.Key.Split('.');

                uniqueFirst.Add(keys[0]);
                if (keys.Length == 2 && !properties.ContainsKey(keys[1]))
                {
                    properties.Add(keys[1], keys[0]);
                }
                else if (keys.Length == 3 && !uniqueSecond.ContainsKey(keys[1]))
                {
                    uniqueSecond.Add(keys[1], keys[0]);
                    if (!properties.ContainsKey(keys[2]))
                        properties.Add(keys[2], keys[1]);
                }
            }
            uniqueFirst = uniqueFirst.Distinct<string>().ToList();

            foreach (string key in uniqueFirst)
            {
                en += $"\"{key}\" : {{{Environment.NewLine}";
                var firstChild = uniqueSecond.Where(a => a.Value == key).Select(s => s.Key).ToList();
                if (firstChild != null && firstChild.Count > 0)
                {
                    foreach (string child in firstChild)
                    {
                        en += $"\"{child}\" : {{{Environment.NewLine}";
                        var firstChildProperties = languages.Where(a => a.Key.Contains(key + "." + child + "."));
                        foreach (var item in firstChildProperties)
                        {
                            en += $"\"{item.Key.Split('.')[2]}\":\"{item.Norwegian}\",{Environment.NewLine}";
                        }
                        en += $"}},{Environment.NewLine}";
                    }

                }
                var children = properties.Where(a => a.Value == key).Select(s => s.Key).ToList();
                if (children != null && children.Count > 0)
                {
                    foreach (string child in children)
                    {
                        var item = languages.FirstOrDefault(a => a.Key.Equals(key + "." + child));
                        if (item != null)
                        {
                            en += $"\"{item.Key.Split('.')[1]}\":\"{item.Norwegian}\",{Environment.NewLine}";
                        }
                    }

                }
                en += $"}},{Environment.NewLine}";
            }

            en += "}";
            File.WriteAllText("Norwegian.json", en, Encoding.UTF8);
        }
       
        private static void ConvertToEnglish(List<Languages> languages)
        {
            List<string> uniqueFirst = new List<string>();
            Dictionary<string, string> uniqueSecond = new Dictionary<string, string>();
            Dictionary<string, string> properties = new Dictionary<string, string>();
            string en = $"{{ {Environment.NewLine}";
            foreach (var item in languages)
            {
                string[] keys = item.Key.Split('.');

                uniqueFirst.Add(keys[0]);
                if (keys.Length == 2 && !properties.ContainsKey(keys[1]))
                {
                    properties.Add(keys[1], keys[0]);
                }
                else if (keys.Length == 3 && !uniqueSecond.ContainsKey(keys[1]))
                {
                    uniqueSecond.Add(keys[1], keys[0]);
                    if (!properties.ContainsKey(keys[2]))
                        properties.Add(keys[2], keys[1]);
                }
            }
            uniqueFirst = uniqueFirst.Distinct<string>().ToList();

            foreach (string key in uniqueFirst)
            {
                en += $"\"{key}\" : {{{Environment.NewLine}";
                var firstChild = uniqueSecond.Where(a => a.Value == key).Select(s => s.Key).ToList();
                if (firstChild != null && firstChild.Count > 0)
                {
                    foreach (string child in firstChild)
                    {
                        en += $"\"{child}\" : {{{Environment.NewLine}";
                        var firstChildProperties = languages.Where(a => a.Key.Contains(key + "." + child + "."));
                        foreach (var item in firstChildProperties)
                        {
                            en += $"\"{item.Key.Split('.')[2]}\":\"{item.English}\",{Environment.NewLine}";
                        }
                        en += $"}},{Environment.NewLine}";
                    }

                }
                var children = properties.Where(a => a.Value == key).Select(s => s.Key).ToList();
                if (children != null && children.Count > 0)
                {
                    foreach (string child in children)
                    {
                        var item = languages.FirstOrDefault(a => a.Key.Equals(key + "." + child));
                        if (item != null)
                        {
                            en += $"\"{item.Key.Split('.')[1]}\":\"{item.English}\",{Environment.NewLine}";
                        }
                    }

                }
                en += $"}},{Environment.NewLine}";
            }

            en += "}";
            File.WriteAllText("English.json", en, Encoding.UTF8);
        }

        public static DataTable ConvertExcelToDataTable(string FileName)
        {
            DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }

        private static void ReadTwocolumnCSV()
        {
            string filePath = @"ServiceKey.csv";
            string[] contents = File.ReadAllLines(filePath);
            string json = @"{" + Environment.NewLine;
            foreach (string content in contents)
            {
                string[] keyValue = content.Split(',');
                json += $"\"{keyValue[0]}\":\"{keyValue[1]}\",{Environment.NewLine}";
            }
            json += "}";
            File.WriteAllText("ServiceKey.json", json, Encoding.UTF8);
        }
    }

    public class Languages
    {
        public string Key { get; set; }
        public string English { get; set; }
        public string Norwegian { get; set; }
        public string Swedish { get; set; }
        public string Danish { get; set; }
    }
}
