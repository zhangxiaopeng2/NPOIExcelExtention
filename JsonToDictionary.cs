using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.ExcelReaderHelper
{
    public class JsonToDictionary
    {
        private static string GetFileJson(string filepath)
        {
            string json = string.Empty;
            using (FileStream fs = new FileStream(filepath, FileMode.Open, System.IO.FileAccess.Read, FileShare.ReadWrite))
            {
                using (StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default))
                {
                    json = sr.ReadToEnd().ToString();
                }
            }
            return json;
        }
        public static Dictionary<string, string> GetDicByJsonFile(string filpath)
        {
            string json = GetFileJson(filpath);
            return JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
        }
    }
}
