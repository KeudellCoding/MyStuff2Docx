using MyStuff2Docx.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace MyStuff2Docx {
    class MyStuffHandler : IDisposable {
        public string BaseZipArchiveTempPath { get; private set; }
        public string TempImagesPath { get; private set; }
        public string TempDocxPath { get; private set; }


        public List<MyStuffCategory> Categories { get; private set; }


        public MyStuffHandler(string path) {
            BaseZipArchiveTempPath = Path.GetTempPath() + "MyStuff2Docx_ZipFiles_" + Guid.NewGuid() + Path.DirectorySeparatorChar;
            TempImagesPath = Path.GetTempPath() + "MyStuff2Docx_Images_" + Guid.NewGuid() + Path.DirectorySeparatorChar;
            TempDocxPath = Path.GetTempPath() + "MyStuff2Docx_DocxFiles_" + Guid.NewGuid() + Path.DirectorySeparatorChar;
            Directory.CreateDirectory(BaseZipArchiveTempPath);
            Directory.CreateDirectory(TempImagesPath);
            Directory.CreateDirectory(TempDocxPath);

            ZipFile.ExtractToDirectory(path, BaseZipArchiveTempPath);

            Categories ??= new List<MyStuffCategory>();
            foreach (var compressedCategory in Directory.GetFiles(BaseZipArchiveTempPath, "*.zip")) {
                var newCategory = new MyStuffCategory() {
                    ZipFileName = Path.GetFileName(compressedCategory),
                    Name = Path.GetFileNameWithoutExtension(compressedCategory),
                    TempImagesPath = TempImagesPath,
                    LocalZipFilePath = compressedCategory
                };

                Categories.Add(newCategory);
            }
        }

        public void Dispose() {
            foreach (var category in Categories ?? new List<MyStuffCategory> { }) {
                category.Dispose();
            }

            Directory.Delete(BaseZipArchiveTempPath, true);
            Directory.Delete(TempImagesPath, true);
            Directory.Delete(TempDocxPath, true);
        }
    }
}
