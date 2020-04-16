using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MyStuff2Docx.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace MyStuff2Docx {
    class WordCreator {
        public const int IMAGE_WIDTH = 990000;

        private List<string> pagePaths = new List<string>();

        public void AddPage(MyStuffItemInfo itemInfo, in MyStuffCategory category, string tempDocxPath) {
            var pagePath = tempDocxPath + "MyStuff2Docx_" + Guid.NewGuid() + ".docx";

            using (var baseDoc = WordprocessingDocument.Create(pagePath, WordprocessingDocumentType.Document)) {
                baseDoc.AddMainDocumentPart();
                baseDoc.MainDocumentPart.Document = new Document();
                baseDoc.MainDocumentPart.Document.Body = new Body();

                #region TableProperties
                var tableProperties = new TableProperties(
                    new TableBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 2 },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 2 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 2 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 2 },
                        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 2 },
                        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 2 }
                    ),
                    new TableWidth() {
                        Width = "5000",
                        Type = TableWidthUnitValues.Pct
                    }
                );
                #endregion

                var table = new Table(tableProperties);

                #region Add CategoryName
                table.AppendChild(new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("Category")))),
                    new TableCell(new Paragraph(new Run(new Text(category.Name))))
                ));
                #endregion

                #region Adding AdditionalProperties
                foreach (var property in itemInfo.AdditionalProperties.Where(itemInfo.AdditionalPropertyFilter)) {
                    table.AppendChild(new TableRow(
                        new TableCell(new Paragraph(new Run(new Text(property.Key)))),
                        new TableCell(new Paragraph(new Run(new Text(property.Value))))
                    ));
                }
                #endregion

                #region Adding Images
                var imagesForThisItem = category.Images.Where(i => itemInfo.Images.Contains(i.PathInCategory));

                var readyImages = new List<Drawing> { };
                foreach (var imageForThisItem in imagesForThisItem) {
                    readyImages.Add(getImageElement(baseDoc, imageForThisItem));
                }

                table.AppendChild(new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("Images")))),
                    new TableCell(new Paragraph(new Run(readyImages)))
                ));
                #endregion


                baseDoc.MainDocumentPart.Document.Body.AppendChild(table);
                baseDoc.MainDocumentPart.Document.Body.AppendChild(new Break() { Type = BreakValues.Page });

                baseDoc.Close();
            }

            pagePaths.Add(pagePath);
        }

        public void CombinePages(string targetFilePath) {
            using (var baseDoc = WordprocessingDocument.Create(targetFilePath, WordprocessingDocumentType.Document)) {
                baseDoc.AddMainDocumentPart();
                baseDoc.MainDocumentPart.Document = new Document();
                baseDoc.MainDocumentPart.Document.Body = new Body();


                for (int i = 0; i < pagePaths.Count; i++) {
                    var chunk = baseDoc.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, "AltChunkId" + i);
                    using (var fileStream = File.Open(pagePaths[i], FileMode.Open)) {
                        chunk.FeedData(fileStream);
                    }
                    var altChunk = new AltChunk() {
                        Id = "AltChunkId" + i
                    };
                    baseDoc.MainDocumentPart.Document.Body.AppendChild(new Break() { Type = BreakValues.Page });
                    baseDoc.MainDocumentPart.Document.Body.AppendChild(altChunk);
                }


                baseDoc.Close();
            }
        }


        private Drawing getImageElement(WordprocessingDocument baseDoc, MyStuffImage image) {
            var imagePart = baseDoc.MainDocumentPart.AddImagePart(image.ImageType);

            if (!getImageDimensions(image.TempImagePath, out int width, out int height)) {
                width = IMAGE_WIDTH;
                height = 792000;
            }
            else {
                var proportionMultiplier = (double)width / IMAGE_WIDTH;
                width = IMAGE_WIDTH;
                height = (int)Math.Ceiling(height / proportionMultiplier);
            }

            using (FileStream stream = new FileStream(image.TempImagePath, FileMode.Open)) {
                imagePart.FeedData(stream);
            }

            // Define the reference of the image.
            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = width, Cy = height },
                    new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties() { Id = 1U, Name = image.ImageId },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties() { Id = 0U, Name = image.ImageFileName },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension() {
                                                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                            }
                                        )
                                    ) {
                                        Embed = baseDoc.MainDocumentPart.GetIdOfPart(imagePart),
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = width, Cy = height }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList()
                                    ) { Preset = A.ShapeTypeValues.Rectangle }))
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                ) {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U,
                    EditId = "50D07946"
                });

            return element;
        }

        private bool getImageDimensions(string path, out int width, out int height) {
            try {
                using (var img = Image.FromFile(path)) {
                    width = img.Width;
                    height = img.Height;
                }
                return true;
            }
            catch (Exception) {
                width = 0;
                height = 0;
                return false;
            }
        }
    }
}
