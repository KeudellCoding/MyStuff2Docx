using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace MyStuff2Docx {
    class Program {
        private const string DEV_BASE_ZIP_PATH = null;
        private const string DEV_TARGET_DIR_PATH = null;

        public static object _CONSOLE_WRITE_LOCK = new object();

        static void Main(string[] args) {
            Console.Title = "MyStuff2Docx";
            MyStuffHandler myStuffHandler = null;

            #region GetMyStuffHandler
            renderHeadline();
            while (myStuffHandler == null) {
                lock (_CONSOLE_WRITE_LOCK) {
                    var zipFilePath = string.Empty;

                    if (string.IsNullOrEmpty(DEV_BASE_ZIP_PATH)) {
                        Console.WriteLine();
                        Console.WriteLine("Please enter the complete path of the exported MyStuff .zip file:");
                        Console.Write("  >> ");
                        zipFilePath = Console.ReadLine();
                    }
                    else
                        zipFilePath = DEV_BASE_ZIP_PATH;
                    
                    zipFilePath = zipFilePath.Trim('"');

                    try {
                        myStuffHandler = new MyStuffHandler(zipFilePath);
                    }
                    catch (FileNotFoundException) {
                        Console.WriteLine("The file could not be found. Please try again.");
                    }
                    catch (Exception) {
                        Console.WriteLine("An unknown error has occurred. Please try again and make sure that you have read access to the file.");
                    }
                }
            }
            #endregion

            using (myStuffHandler) {

                #region GetSelection
                var selectedPage = 1;
                var maxPages = Math.Ceiling((double)myStuffHandler.Categories.Count / 9);

                do {
                    lock (_CONSOLE_WRITE_LOCK) {
                        renderHeadline();

                        var index = 1;
                        foreach (var category in myStuffHandler.Categories.Skip(9 * (selectedPage - 1)).Take(9)) {
                            Console.Write(category.Selected ? "[X] " : "[ ] ");
                            Console.Write($"{index} {category.Name}");
                            Console.WriteLine();

                            index++;
                        }
                        Console.WriteLine();
                        Console.WriteLine($" -- Page {selectedPage}/{maxPages} --");
                        Console.WriteLine();
                        Console.WriteLine(" (1-9 -> select | w -> up | s -> down | a -> toogle all | x -> cancel | b -> save)");
                        Console.WriteLine();
                        Console.Write("  >> ");
                        var rawInput = Console.ReadKey();

                        if (rawInput.Key == ConsoleKey.W && selectedPage < maxPages)
                            selectedPage++;
                        else if (rawInput.Key == ConsoleKey.S && selectedPage > 1)
                            selectedPage--;
                        else if (rawInput.Key == ConsoleKey.A) {
                            var newValue = !myStuffHandler.Categories?.FirstOrDefault()?.Selected ?? false;
                            myStuffHandler.Categories.ForEach(cat => cat.Selected = newValue);
                        }
                        else if (int.TryParse(rawInput.KeyChar.ToString(), out int selectedOption) && selectedOption > 0 && selectedOption <= 9) {
                            var selectedItem = myStuffHandler.Categories.ElementAtOrDefault(selectedOption + (9 * (selectedPage - 1)) - 1);
                            if (selectedItem != null) {
                                selectedItem.Selected = !selectedItem.Selected;
                            }
                        }
                        else if (rawInput.Key == ConsoleKey.X)
                            return;
                        else if (rawInput.Key == ConsoleKey.B)
                            break;
                    }
                } while (true);
                #endregion

                #region GetFolderForExportAndExportType
                renderHeadline();
                var exportFolderPath = string.Empty;
                var multipleFiles = true;

                do {
                    lock (_CONSOLE_WRITE_LOCK) {
                        if (string.IsNullOrEmpty(DEV_TARGET_DIR_PATH)) {
                            Console.WriteLine();
                            Console.WriteLine("Please enter the complete path for the created documents:");
                            Console.Write("  >> ");
                            exportFolderPath = Console.ReadLine();
                        }
                        else
                            exportFolderPath = DEV_TARGET_DIR_PATH;

                        exportFolderPath = exportFolderPath.Trim('"');
                        exportFolderPath = exportFolderPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

                        if (!string.IsNullOrWhiteSpace(exportFolderPath) && Directory.Exists(exportFolderPath))
                            break;
                        else
                            Console.WriteLine("This is not a valid directory. Please try again.");
                    }
                } while (true);

                do {
                    lock (_CONSOLE_WRITE_LOCK) {
                        Console.WriteLine();
                        Console.WriteLine("Would you like to have one document for all categories (not recommended),");
                        Console.WriteLine("or rather one document per category (recommended)?");
                        Console.WriteLine("Please note: Word documents that are larger than 500 mb");
                        Console.WriteLine("or have more than 1000 pages are sometimes difficult or impossible to open.");
                        Console.WriteLine(" (o -> one document | m -> many documents)");
                        Console.Write("  >> ");
                        var rawInput = Console.ReadKey();

                        if (rawInput.Key == ConsoleKey.O) {
                            multipleFiles = false;
                            break;
                        }
                        else if (rawInput.Key == ConsoleKey.M) {
                            multipleFiles = true;
                            break;
                        }
                    }
                } while (true);
                #endregion

                #region ExportCategories
                renderHeadline();
                if (multipleFiles) {
                    var processingMessages = new Dictionary<string, ProcessingMessage> { };
                    var generationTasks = new List<Task> { };

                    foreach (var selectedCategory in myStuffHandler.Categories.Where(cat => !cat.Selected)) {
                        selectedCategory.Dispose();
                    }

                    foreach (var selectedCategory in myStuffHandler.Categories.Where(cat => cat.Selected)) {
                        processingMessages.Add($"{selectedCategory.Name}_GeneratingAndCombiningPages", new ProcessingMessage($"Generating and combining \"{selectedCategory.Name}\""));
                        
                        var newTask = Task.Factory.StartNew(() => {
                            var wordCreator = new WordCreator();
                            var exportFilePath = @$"{exportFolderPath}\MyStuff2Docx_{selectedCategory.Name}_Export_{Guid.NewGuid()}.docx";

                            try {
                                foreach (var itemInfo in selectedCategory.ItemInfos) {
                                    wordCreator.AddPage(itemInfo, selectedCategory, myStuffHandler.TempDocxPath);
                                }

                                wordCreator.CombinePages(exportFilePath);
                            }
                            catch (Exception) { processingMessages[$"{selectedCategory.Name}_GeneratingAndCombiningPages"].HasError = true; }
                            finally { processingMessages[$"{selectedCategory.Name}_GeneratingAndCombiningPages"].Dispose(); }

                            processingMessages.Remove($"{selectedCategory.Name}_GeneratingAndCombiningPages");
                            selectedCategory.Dispose();
                        });

                        generationTasks.Add(newTask);
                    }

                    Task.WaitAll(generationTasks.ToArray());
                }
                else {
                    var wordCreator = new WordCreator();
                    var exportFilePath = @$"{exportFolderPath}\MyStuff2Docx_Export_{Guid.NewGuid()}.docx";

                    foreach (var selectedCategory in myStuffHandler.Categories.Where(cat => cat.Selected)) {
                        using (var pm = new ProcessingMessage($"\"{selectedCategory.Name}\"")) {
                            try { 
                                foreach (var itemInfo in selectedCategory.ItemInfos) {
                                    wordCreator.AddPage(itemInfo, selectedCategory, myStuffHandler.TempDocxPath);
                                }
                            }
                            catch (Exception) { pm.HasError = true; }
                        }

                        selectedCategory.Dispose();
                    }

                    Console.WriteLine();
                    using (var pm = new ProcessingMessage($"Combining all pages")) {
                        try {
                            wordCreator.CombinePages(exportFilePath);
                        }
                        catch (Exception) { pm.HasError = true; }
                    }
                }
                #endregion

                GC.Collect();
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("The program has run through. Please press any button to exit.");
                Console.ReadKey();
            }
        }


        private static void renderHeadline(bool clearBefore = true) {
            lock (_CONSOLE_WRITE_LOCK) {
                if (clearBefore) {
                    Console.Clear();
                }

                Console.WriteLine("==================================");
                Console.WriteLine("========== MyStuff2Docx ==========");
                Console.WriteLine("==================================");
                Console.WriteLine();
            }
        }
    }
}
