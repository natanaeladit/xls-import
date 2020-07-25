using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;

namespace ImportConsole
{
    public class Branch
    {
        public string Country { get; set; }
        public string BankName { get; set; }
        public string BankCode { get; set; }
        public string BankBranchName { get; set; }
        public string BranchCode { get; set; }
        public string CodeType1 { get; set; }
        public string Code1 { get; set; }
        public string CodeType2 { get; set; }
        public string Code2 { get; set; }
        public string SyntaxRestrictionOnAccountNo { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Insert filename (without .xlsx extension): ");
            string filename = Console.ReadLine();
            filename += ".xlsx";

            if (File.Exists(filename))
            {
                List<Branch> branches = new List<Branch>();

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(filename, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        bool skipHeader = reader.Read();
                        while (reader.Read())
                        {
                            branches.Add(new Branch()
                            {
                                Country = reader.GetValue(1)?.ToString(),
                                BankName = reader.GetValue(2)?.ToString(),
                                BankCode = reader.GetValue(3)?.ToString(),
                                BankBranchName = reader.GetValue(4)?.ToString(),
                                BranchCode = reader.GetValue(5)?.ToString(),
                                CodeType1 = reader.GetValue(6)?.ToString(),
                                Code1 = reader.GetValue(7)?.ToString(),
                                CodeType2 = reader.GetValue(8)?.ToString(),
                                Code2 = reader.GetValue(9)?.ToString(),
                                SyntaxRestrictionOnAccountNo = reader.GetValue(10)?.ToString(),
                            });
                        }
                    }
                }
            }
            else
                Console.WriteLine($"Cannot find file {filename}");
        }
    }
}
