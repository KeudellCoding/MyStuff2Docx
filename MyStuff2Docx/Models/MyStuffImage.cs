using DocumentFormat.OpenXml.Packaging;

namespace MyStuff2Docx.Models {
    class MyStuffImage {
        public string PathInCategory { get; set; }
        public string ImageFileName { get; set; }
        public string ItemId { get; set; }
        public string ImageId { get; set; }
        public ImagePartType ImageType { get; set; }

        public string TempImagePath { get; set; }
    }
}
