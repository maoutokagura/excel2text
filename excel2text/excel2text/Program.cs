using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using OfficeOpenXml;
using ClosedXML.Excel;

namespace excel2text {
    class Program {
        static void Main(string[] args) {
            Console.WriteLine("start extract");
            string current_path = System.IO.Directory.GetCurrentDirectory();
            string[] workbook_paths = Directory.GetFiles(current_path + @"\bin", "*");

            string common_output_path = current_path + "\\src\\";
            SafeCreateDirectory(common_output_path);

            //            try {
            foreach (var workbook_path in workbook_paths) {
                string book_output_path = common_output_path + System.IO.Path.GetFileNameWithoutExtension(workbook_path);

                SafeCreateDirectory(book_output_path);
                OutputNames(workbook_path, book_output_path);

                XLWorkbook workbook = new XLWorkbook(workbook_path);
                IXLWorksheets sheets = workbook.Worksheets;

                Encoding Enc = Encoding.GetEncoding("UTF-8");

                foreach (var sheet in sheets) {
                    string safe_sheet_name = Regex.Replace(sheet.Name, @"[^\w\.@-]", "", RegexOptions.None);

                    var value_direname = book_output_path + "\\value\\";
                    SafeCreateDirectory(value_direname);
                    using (StreamWriter writer = new StreamWriter(value_direname + safe_sheet_name + ".txt", false, Enc)) {
                        foreach (var item in sheet.CellsUsed()) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item.Value);
                        }
                    }

                    var style_dirname = book_output_path + "\\style\\";
                    SafeCreateDirectory(style_dirname);
                    using (StreamWriter writer = new StreamWriter(style_dirname + safe_sheet_name + ".txt", false, Enc)) {
                        foreach (var item in sheet.CellsUsed()) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item.Style);
                        }
                    }

                    var FormulaA1_dirname = book_output_path + "\\A1\\";
                    SafeCreateDirectory(FormulaA1_dirname);
                    using (StreamWriter writer = new StreamWriter(FormulaA1_dirname + safe_sheet_name + ".txt", false, Enc)) {
                        foreach (var item in sheet.CellsUsed()) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item.FormulaA1);
                        }
                    }

                    var FormulaR1C1_dirname = book_output_path + "\\R1C1\\";
                    SafeCreateDirectory(FormulaR1C1_dirname);
                    using (StreamWriter writer = new StreamWriter(FormulaR1C1_dirname + safe_sheet_name + ".txt", false, Enc)) {
                        foreach (var item in sheet.CellsUsed()) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item.FormulaR1C1);
                        }
                    }

                    var item_filename = book_output_path + "\\item\\";
                    SafeCreateDirectory(item_filename);
                    using (StreamWriter writer = new StreamWriter(item_filename + safe_sheet_name + ".txt", false, Enc)) {
                        foreach (var item in sheet.CellsUsed()) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item);
                        }
                    }
                }
            }
            //} catch (Exception e) {
            //    Console.WriteLine("{0} Exception caught.", e);
            //}
        }

        public static void OutputNames(string input_book_path, string output_path) {
            FileInfo workbook_info = new FileInfo(input_book_path);
            ExcelPackage package = new ExcelPackage(workbook_info);
            ExcelWorkbook workbook = package.Workbook;
            ExcelWorksheets sheets = workbook.Worksheets;

            Encoding Enc = Encoding.GetEncoding("UTF-8");
            string file_name = Path.GetFileNameWithoutExtension(input_book_path);

            using (StreamWriter writer = new StreamWriter(output_path + @"\" + file_name + "_names.txt", false, Enc)) {
                foreach (var item in workbook.Names) {
                    writer.WriteLine(item.Name + "\t" + item.Address);
                }
            }
            return;
        }


        public static DirectoryInfo SafeCreateDirectory(string path) {
            if (Directory.Exists(path)) {
                return null;
            }
            return Directory.CreateDirectory(path);
        }
    }


}
