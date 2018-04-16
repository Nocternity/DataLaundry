using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;

namespace DataLaundry
{
    public class JsonHandler
    {
        public void JsonFixer(List<Company> company)
        {
            string filename = @"d:\companies.json";

            File.WriteAllText(filename, JsonConvert.SerializeObject(company));

            // serialize JSON directly to a file
            using (StreamWriter file = File.AppendText(filename))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, company);

            }


        }
    }
}
