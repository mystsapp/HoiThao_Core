//using System;
//using System.Collections.Generic;
//using System.Data.SqlClient;
//using System.IO;
//using System.Linq;
//using System.Threading.Tasks;
//using Microsoft.Extensions.Configuration;
//using FastMember;

//namespace BulkCopy
//{
//    public class Program
//    {
//        public static IConfigurationRoot Configuration;
//        public static void Main(string[] args)
//        {
//            var builder = new ConfigurationBuilder()
//               .SetBasePath(Directory.GetCurrentDirectory())
//               .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
//               .AddEnvironmentVariables();

//            builder.AddEnvironmentVariables();
//            Configuration = builder.Build();
//            string connectionstring = Configuration["ConnectionString"];

//            List<string> records = new List<string>();
//            using (StreamReader sr = new StreamReader(File.OpenRead("C:\\Resources\\Employee.txt")))
//            {

//                string file = sr.ReadToEnd();
//                records = new List<string>(file.Split('\n'));
//            }


//            List<Employee> emplist = new List<Employee>();

//            foreach (string record in records)
//            {
//                Employee emp = new Employee();
//                string[] textpart = record.Split('|');
//                emp.EmpId = textpart[0];
//                emp.EmpName = textpart[1];
//                emp.Salary = Convert.ToDecimal(textpart[2]);
//                emplist.Add(emp);

//            }


//            var copyParameters = new[]
//             {
//                        nameof(Employee.EmpId),
//                        nameof(Employee.EmpName),
//                        nameof(Employee.Salary)

//                    };


//            using (var sqlCopy = new SqlBulkCopy(connectionstring))
//            {
//                sqlCopy.DestinationTableName = "[Employee]";
//                sqlCopy.BatchSize = 500;
//                using (var reader = ObjectReader.Create(emplist, copyParameters))
//                {
//                    sqlCopy.WriteToServer(reader);
//                }
//            }


//        }
//    }

//    class Employee
//    {
//        public string EmpId { get; set; }
//        public string EmpName { get; set; }
//        public decimal Salary { get; set; }
//    }

//}
//The package.json
//{  
//  "version": "1.0.0-*",  
//  "buildOptions": {
//        "emitEntryPoint": true
//  },  
  
//  "dependencies": {
//        "Microsoft.Extensions.Configuration": "1.1.0",  
//    "Microsoft.Extensions.Configuration.FileExtensions": "1.1.0",  
//    "Microsoft.Extensions.Configuration.Json": "1.1.0",  
//    "Microsoft.NETCore.App": {
//            "type": "platform",  
//      "version": "1.0.1"
//    },  
//    "System.IO.FileSystem": "4.3.0",  
//    "Microsoft.Extensions.Configuration.EnvironmentVariables": "1.1.0",  
//    "System.Data.SqlClient": "4.1.0",  
//    "FastMember": "1.1.0"
//  },  
  
//  "frameworks": {
//        "netcoreapp1.0": {
//            "imports": "dnxcore50"
//        }
//    }
//}