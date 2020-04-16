using CsvHelper;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace MyStuff2Docx.Models {
    class MyStuffCategory : IDisposable {
        private ZipArchive categoryZipArchive;
        public ZipArchive CategoryZipArchive => categoryZipArchive ??= ZipFile.OpenRead(LocalZipFilePath);


        public string TempImagesPath { get; set; }
        public string LocalZipFilePath { get; set; }
        public string ZipFileName { get; set; }
        public string Name { get; set; }
        public bool Selected { get; set; } = false;


        private List<MyStuffItemInfo> itemInfos;
        public List<MyStuffItemInfo> ItemInfos => itemInfos ??= getItemInfos();

        private List<MyStuffImage> images;
        public List<MyStuffImage> Images => images ??= getImages();



        private List<MyStuffItemInfo> getItemInfos() {
            var result = new List<MyStuffItemInfo>();

            var infoCsvFile = CategoryZipArchive.Entries.SingleOrDefault(e => e.Name.Equals($"{Name}.csv"));
            if (infoCsvFile != null) {
                using (var rawCsvReader = new StreamReader(infoCsvFile.Open())) {
                    using (var csv = new CsvReader(rawCsvReader, CultureInfo.InvariantCulture)) {
                        csv.Read();
                        csv.ReadHeader();
                        while (csv.Read()) {
                            var newItemInfo = new MyStuffItemInfo() {
                                Id = csv.GetField<string>("item id"),
                                ItemLocation = csv.GetField<string>("item location"),
                                Images = csv.GetField<string>("item images").Split("|").Where(x => !string.IsNullOrEmpty(x)).ToArray(),
                                Attachments = csv.GetField<string>("item attachments").Split("|").Where(x => !string.IsNullOrEmpty(x)).ToArray()
                            };

                            try {
                                newItemInfo.ItemUpdated = DateTime.ParseExact(csv.GetField<string>("item updated"), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            }
                            catch (Exception) { }
                            try {
                                newItemInfo.ItemCreated = DateTime.ParseExact(csv.GetField<string>("item created"), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            }
                            catch (Exception) { }

                            newItemInfo.AdditionalProperties ??= new Dictionary<string, string> { };
                            foreach (var header in csv.Context.HeaderRecord) {
                                newItemInfo.AdditionalProperties.TryAdd(header, csv.TryGetField(header, out string value) ? value : string.Empty);
                            }

                            result.Add(newItemInfo);
                        }
                    }
                }
            }

            return result;
        }
        private List<MyStuffImage> getImages() {
            var result = new List<MyStuffImage>();
            
            foreach (var compressedImage in CategoryZipArchive.Entries) {
                var imageInfo = new MyStuffImage();

                switch (Path.GetExtension(compressedImage.Name)) {
                    case ".jpg":
                    case ".jpeg":
                        imageInfo.ImageType = ImagePartType.Jpeg;
                        break;
                    case ".png":
                        imageInfo.ImageType = ImagePartType.Png;
                        break;
                    case ".bmp":
                        imageInfo.ImageType = ImagePartType.Bmp;
                        break;
                    case ".gif":
                        imageInfo.ImageType = ImagePartType.Gif;
                        break;
                    case ".tiff":
                        imageInfo.ImageType = ImagePartType.Tiff;
                        break;
                    default:
                        continue;
                }

                imageInfo.ImageFileName = compressedImage.Name;
                imageInfo.PathInCategory = compressedImage.FullName;
                imageInfo.ImageId = Path.GetFileNameWithoutExtension(compressedImage.Name);
                imageInfo.ItemId = compressedImage.FullName.Split('\\', '/').FirstOrDefault();
                imageInfo.TempImagePath = TempImagesPath + "Temp_Image_" + Guid.NewGuid() + ".tmp";

                using (var compressedImageStream = compressedImage.Open()) {
                    using (var tempImageStream = File.Open(imageInfo.TempImagePath, FileMode.OpenOrCreate)) {
                        compressedImageStream.CopyTo(tempImageStream);
                    }
                }

                result.Add(imageInfo);
            }


            return result;
        }


        public void Dispose() {
            if (categoryZipArchive != null) {
                categoryZipArchive.Dispose();
                categoryZipArchive = null;
            }
        }
    }
}
