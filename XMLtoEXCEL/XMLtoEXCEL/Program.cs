using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace XMLtoEXCEL
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("1.xml");
                XmlElement xmlRoot = xmlDoc.DocumentElement;
                foreach (XmlNode xnode in xmlRoot)
                {
                    string ts = xnode.Attributes[0].Value;
                    Console.WriteLine(ts);
                }

                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "TestLT.xlsx")))
                    {
                        helper.Set(column: "A", row: 1, data: "Crycry");
                        var val = helper.Get(column: "A", row: 6);
                        helper.Set(column: "B", row: 1, data: DateTime.Now);
                        helper.Set(column: "A", row: 3, data: "Crycry");

                        helper.Save();
                    }
                }

                Console.Read();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
