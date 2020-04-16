using System;
using System.Collections.Generic;
using System.Linq;

namespace MyStuff2Docx.Models {
    class MyStuffItemInfo {
        public string Id { get; set; }
        public string ItemLocation { get; set; }

        public Dictionary<string, string> AdditionalProperties { get; set; }

        public string[] Images { get; set; }
        public string[] Attachments { get; set; }
        public DateTime ItemUpdated { get; set; }
        public DateTime ItemCreated { get; set; }

        public bool AdditionalPropertyFilter(KeyValuePair<string, string> property) {
            var conditions = new List<bool> {
                !property.Key.Equals("item id"),
                !property.Key.Equals("item barcode"),
                !property.Key.Equals("item images"),
                !property.Key.Equals("item attachments")
            };

            return conditions.All(c => c);
        }
    }
}
