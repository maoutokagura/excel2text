using System;
using ClosedXML.Excel;
using System.IO;
using System.Collections.Generic;

using System.Text;

namespace ConsoleApp2 {
    class Program {
        static void Main(string[] args) {
            Console.WriteLine("start extract");
            string current_path = System.IO.Directory.GetCurrentDirectory();
            string[] workbook_paths = Directory.GetFiles(current_path + @"\bin", "*");

            try {
                foreach (var workbook_path in workbook_paths) {
                    XLWorkbook workbook = new XLWorkbook(workbook_path);
                    IXLWorksheets sheets = workbook.Worksheets;
                    string common_output_path = current_path + "\\src\\" + System.IO.Path.GetFileNameWithoutExtension(workbook_path);

                    Encoding Enc = Encoding.GetEncoding("UTF-8");

                    foreach (var sheet in sheets) {
                        var value_direname = common_output_path + "\\value\\";
                        SafeCreateDirectory(value_direname);
                        using (StreamWriter writer = new StreamWriter(value_direname + sheet.Name + ".txt", false, Enc)) {
                            foreach (var item in sheet.CellsUsed()) {
                                writer.WriteLine(item.Address.ToString() + "\t" + item.Value);
                            }
                        }

                        var style_dirname = common_output_path + "\\style\\";
                        SafeCreateDirectory(style_dirname);
                        using (StreamWriter writer = new StreamWriter(style_dirname + sheet.Name + ".txt", false, Enc)) {
                            foreach (var item in sheet.CellsUsed()) {
                                writer.WriteLine(item.Address.ToString() + "\t" + item.Style);
                            }
                        }

                        var FormulaA1_dirname = common_output_path + "\\A1\\";
                        SafeCreateDirectory(FormulaA1_dirname);
                        using (StreamWriter writer = new StreamWriter(FormulaA1_dirname + sheet.Name + ".txt", false, Enc)) {
                            foreach (var item in sheet.CellsUsed()) {
                                writer.WriteLine(item.Address.ToString() + "\t" + item.FormulaA1);
                            }
                        }

                        var FormulaR1C1_dirname = common_output_path + "\\R1C1\\";
                        SafeCreateDirectory(FormulaR1C1_dirname);
                        using (StreamWriter writer = new StreamWriter(FormulaR1C1_dirname + sheet.Name + ".txt", false, Enc)) {
                            foreach (var item in sheet.CellsUsed()) {
                                writer.WriteLine(item.Address.ToString() + "\t" + item.FormulaR1C1);
                            }
                        }

                        var item_filename = common_output_path + "\\item\\";
                        SafeCreateDirectory(item_filename);
                        using (StreamWriter writer = new StreamWriter(item_filename + sheet.Name + ".txt", false, Enc)) {
                            foreach (var item in sheet.CellsUsed()) {
                                writer.WriteLine(item.Address.ToString() + "\t" + item);
                            }
                        }
                    }
                    workbook.Save();
                }
            } catch (Exception e) {
                Console.WriteLine("{0} Exception caught.", e);
            }
        }

        public static DirectoryInfo SafeCreateDirectory(string path) {
            if (Directory.Exists(path)) {
                return null;
            }
            return Directory.CreateDirectory(path);
        }
    }
}
