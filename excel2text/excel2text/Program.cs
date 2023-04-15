using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using OfficeOpenXml;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;

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
                try {
                    ClosedXmlOut(book_output_path, workbook_path);
                } catch {
                    EEplusOut(book_output_path, workbook_path);
                }

            }
        }
        public static void EEplusOut(string book_output_path, string workbook_path) {

            FileInfo fileInfo = new FileInfo(workbook_path);
            ExcelPackage excl = new ExcelPackage(fileInfo);

            ExcelWorksheets sheets = excl.Workbook.Worksheets;

            Encoding Enc = Encoding.GetEncoding("UTF-8");

            foreach (var sheet in sheets) {

                string safe_sheet_name = sheet.Name;
                foreach (char c in Path.GetInvalidFileNameChars()) {
                    safe_sheet_name = safe_sheet_name.Replace(c, '_');
                }

                var value_direname = book_output_path + "\\value\\";
                SafeCreateDirectory(value_direname);
                using (StreamWriter writer = new StreamWriter(value_direname + safe_sheet_name + ".txt", false, Enc)) {
                    foreach (var item in sheet.Cells) {
                        string safe_item_value = "";
                        if ((item.Value != null) && (item.Value.ToString() != "")) {
                            safe_item_value = Regex.Replace(item.Value.ToString(), @"[\r\n]", "\\n", RegexOptions.None);
                            writer.WriteLine(item.Address.ToString() + "\t" + safe_item_value);
                        }
                    }
                }

                var style_dirname = book_output_path + "\\style\\";
                SafeCreateDirectory(style_dirname);
                using (StreamWriter writer = new StreamWriter(style_dirname + safe_sheet_name + ".txt", false, Enc)) {
                    foreach (var item in sheet.Cells) {
                        if ((item.StyleName != null) && (item.StyleName.ToString() != "")) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item.StyleName + " " + item.Style.GetHashCode());
                        }
                    }
                }

                var FormulaA1_dirname = book_output_path + "\\A1\\";
                SafeCreateDirectory(FormulaA1_dirname);
                using (StreamWriter writer = new StreamWriter(FormulaA1_dirname + safe_sheet_name + ".txt", false, Enc)) {
                    foreach (var item in sheet.Cells) {
                        if ((item.Formula != null) && (item.Formula.ToString() != "")) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item.Formula);
                        }
                    }
                }

                var FormulaR1C1_dirname = book_output_path + "\\R1C1\\";
                SafeCreateDirectory(FormulaR1C1_dirname);
                using (StreamWriter writer = new StreamWriter(FormulaR1C1_dirname + safe_sheet_name + ".txt", false, Enc)) {
                    foreach (var item in sheet.Cells) {
                        if ((item.FormulaR1C1 != null) && (item.FormulaR1C1.ToString() != "")) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item.FormulaR1C1);
                        }
                    }
                }

                var item_filename = book_output_path + "\\item\\";
                SafeCreateDirectory(item_filename);
                using (StreamWriter writer = new StreamWriter(item_filename + safe_sheet_name + ".txt", false, Enc)) {
                    foreach (var item in sheet.Cells) {
                        if ((item != null) && (item.ToString() != "")) {
                            writer.WriteLine(item.Address.ToString() + "\t" + item);
                        }
                    }
                }
            }
        }
        public static void ClosedXmlOut(string book_output_path, string workbook_path) {

            XLWorkbook workbook = new XLWorkbook(workbook_path);
            IXLWorksheets sheets = workbook.Worksheets;

            Encoding Enc = Encoding.GetEncoding("UTF-8");

            foreach (var sheet in sheets) {

                string safe_sheet_name = sheet.Name;
                foreach (char c in Path.GetInvalidFileNameChars()) {
                    safe_sheet_name = safe_sheet_name.Replace(c, '_');
                }

                var value_direname = book_output_path + "\\value\\";
                SafeCreateDirectory(value_direname);
                using (StreamWriter writer = new StreamWriter(value_direname + safe_sheet_name + ".txt", false, Enc)) {
                    foreach (var item in sheet.CellsUsed()) {
                        string safe_item_value = "";
                        try {
                            safe_item_value = Regex.Replace(item.Value.ToString(), @"[\r\n]", "\\n", RegexOptions.None);
                        } catch {
                            continue;
                        }
                        writer.WriteLine(item.Address.ToString() + "\t" + safe_item_value);
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


        public static void OutputNames(string input_book_path, string output_path) {
            FileInfo workbook_info = new FileInfo(input_book_path);
            ExcelPackage package = new ExcelPackage(workbook_info);
            ExcelWorkbook workbook = package.Workbook;
            ExcelWorksheets sheets = workbook.Worksheets;

            Encoding Enc = Encoding.GetEncoding("UTF-8");
            string file_name = Path.GetFileNameWithoutExtension(input_book_path);

            using (StreamWriter writer = new StreamWriter(output_path + @"\" + file_name + "_names.txt", false, Enc)) {
                foreach (var item in workbook.Names) {
                    Dictionary<string, string> a = new Dictionary<string, string>() { { "FullAddress", "" }, { "Name", "" }, { "A1", "" }, { "R1C1", "" } };
                    a["FullAddress"] = item.FullAddressAbsolute;
                    a["Name"] = item.Name;
                    a["A1"] = item.Formula;
                    writer.WriteLine(a["FullAddress"] + "\t" + a["Name"] + "\t" + a["A1"] + "\t" + a["R1C1"]);
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
