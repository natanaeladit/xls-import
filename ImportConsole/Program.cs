using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
                Console.WriteLine($"Reading file completed: {branches.Count} rows");
                Console.WriteLine($"BankName: {branches.OrderByDescending(e => e.BankName?.Length).FirstOrDefault()?.BankName?.Length}");
                Console.WriteLine($"BankCode: {branches.OrderByDescending(e => e.BankCode?.Length).FirstOrDefault()?.BankCode?.Length}");
                Console.WriteLine($"BankBranchName: {branches.OrderByDescending(e => e.BankBranchName?.Length).FirstOrDefault()?.BankBranchName?.Length}");
                Console.WriteLine($"BranchCode: {branches.OrderByDescending(e => e.BranchCode?.Length).FirstOrDefault()?.BranchCode?.Length}");
                Console.WriteLine($"CodeType1: {branches.OrderByDescending(e => e.CodeType1?.Length).FirstOrDefault()?.CodeType1?.Length}");
                Console.WriteLine($"Code1: {branches.OrderByDescending(e => e.Code1?.Length).FirstOrDefault()?.Code1?.Length}");
                Console.WriteLine($"CodeType2: {branches.OrderByDescending(e => e.CodeType2?.Length).FirstOrDefault()?.CodeType2?.Length}");
                Console.WriteLine($"Code2: {branches.OrderByDescending(e => e.Code2?.Length).FirstOrDefault()?.Code2?.Length}");
                Console.WriteLine($"SyntaxRestrictionOnAccountNo: {branches.OrderByDescending(e => e.SyntaxRestrictionOnAccountNo?.Length).FirstOrDefault()?.SyntaxRestrictionOnAccountNo?.Length}");
            }
            else
                Console.WriteLine($"Cannot find file {filename}");
        }
    }
}
