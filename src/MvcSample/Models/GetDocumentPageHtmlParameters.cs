﻿
namespace MvcSample.Models
{
    public class GetDocumentPageHtmlParameters
    {
        public string Path { get; set; }

        public int PageIndex { get; set; }

        public bool UsePngImages { get; set; }

        public bool EmbedImagesIntoHtmlForWordFiles { get; set; }

        public string InstanceIdToken { get; set; }

        public string Locale { get; set; }

        public bool SaveFontsInAllFormats { get; set; }
    }
}