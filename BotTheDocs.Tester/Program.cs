using Autofac;
using Newtonsoft.Json.Linq;
using SupportSearch.Common.Interfaces;
using SupportSearch.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTester
{
    class Program
    {
        private static IContainer Container { get; set; }

        static void Main(string[] args)
        {
            Init();

            string docPath = "C:\\Temp\\MicrosotFlowBasics.docx";

            if (System.IO.File.Exists(docPath))
            {
                try
                {
                    Console.WriteLine($"Converting document {docPath}");
                    System.IO.Stream strm = System.IO.File.OpenRead(docPath);
                    SupportSearch.Helper.WordConverterService converter = Container.Resolve<WordConverterService>();
                    converter.ConvertDocToHtml("Microsoft Flow Overview", strm);
                    Console.WriteLine("Conversion finished! Press any key to quit..");
                    Console.ReadLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error converting document:");
                    Console.WriteLine(ex.ToString());
                    Console.ReadLine();
                }
            }
            else
            {
                Console.WriteLine("Input document does not exist!");
                Console.ReadLine();
            }
        }

        private static void Init()
        {
            Newtonsoft.Json.Linq.JObject config = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText("local.settings.json"));

            if (config["Values"] != null)
            {
                foreach (JProperty prop in config["Values"])
                {
                    System.Environment.SetEnvironmentVariable(prop.Name, prop.Value.ToString());
                }
            }

            var builder = new ContainerBuilder();

            // Register individual components
            builder.RegisterInstance(new FileUploadHelper())
                   .As<ISupportBlobUpload>();
            builder.RegisterInstance(new BasicKeyWordsHelper())
                   .As<ISupportKeyWords>();
            builder.RegisterType<WordConverterService>();
            Container = builder.Build();
        }
    }
}
