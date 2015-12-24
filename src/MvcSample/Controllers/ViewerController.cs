using GroupDocs.Viewer.Config;
using GroupDocs.Viewer.Converter.Option;
using GroupDocs.Viewer.Domain;
using GroupDocs.Viewer.Domain.Requests;
using GroupDocs.Viewer.Handler;
using MvcSample.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using WatermarkPosition = MvcSample.Models.WatermarkPosition;

namespace MvcSample.Controllers
{
    public class ViewerController : Controller
    {
        private readonly ViewerHtmlHandler _htmlHandler;
        private readonly ViewerImageHandler _imageHandler;

        private string _licensePath = "D:\\vlitvinchik\\sites\\yanalitvinchik.com\\GroupDocs.Viewer.lic";
        private string _storagePath = AppDomain.CurrentDomain.GetData("DataDirectory").ToString(); // App_Data folder path
        private string _tempPath = AppDomain.CurrentDomain.GetData("DataDirectory") + "\\Temp";
        private readonly ViewerConfig _config;

        public ViewerController()
        {
            _config = new ViewerConfig
            {
                StoragePath = _storagePath,
                TempPath = _tempPath,
                UseCache = true
            };

            _htmlHandler = new ViewerHtmlHandler(_config);
            _imageHandler = new ViewerImageHandler(_config);

            _htmlHandler.SetLicense(_licensePath);
            _imageHandler.SetLicense(_licensePath);
        }

        // GET: /Viewer/
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ViewDocument(ViewDocumentParameters request)
        {
            var result = new ViewDocumentResponse
            {
                pageCss = new string[] { },
                lic = true
            };

            if (request.UseHtmlBasedEngine)
            {
                var docInfo = _htmlHandler.GetDocumentInfo(new GetDocumentInfoRequest(request.Path));
                result.documentDescription = new FileDataJsonSerializer(docInfo.FileData, new FileDataOptions()).Serialize(false);
                result.docType = docInfo.DocumentType;
                result.fileType = docInfo.DocumentFileType;

                var htmlOptions = new HtmlOptions { IsResourcesEmbedded = true, Watermark = GetWatermark(request) };
                var htmlPages = _htmlHandler.GetPages(new FileDescription { Guid = request.Path, Name = request.Path }, htmlOptions);
                result.pageHtml = htmlPages.Select(_ => _.HtmlContent).ToArray();

                //NOTE: Fix for incomplete cells document
                for (int i = 0; i < result.pageHtml.Length; i++)
                {
                    var html = result.pageHtml[i];
                    var indexOfScript = html.IndexOf("script");
                    if (indexOfScript > 0)
                        result.pageHtml[i] = html.Substring(0, indexOfScript);
                }
            }
            else
            {
                var docInfo = _imageHandler.GetDocumentInfo(new GetDocumentInfoRequest(request.Path));
                result.documentDescription = new FileDataJsonSerializer(docInfo.FileData, new FileDataOptions()).Serialize(true);
                result.docType = docInfo.DocumentType;
                result.fileType = docInfo.DocumentFileType;

                var imageOptions = new ImageOptions { Watermark = GetWatermark(request) };
                var imagePages = _imageHandler.GetPages(new FileDescription { Guid = request.Path, Name = request.Path }, imageOptions);

                // Provide images urls
                var urls = new List<string>();

                // If no cache - save images to temp folder
                var tempFolderPath = Path.Combine(Server.MapPath("~"), "Content", "TempStorage");

                foreach (var pageImage in imagePages)
                {
                    var docFoldePath = Path.Combine(tempFolderPath, request.Path);

                    if (!Directory.Exists(docFoldePath))
                        Directory.CreateDirectory(docFoldePath);

                    var pageImageName = string.Format("{0}\\{1}.png", docFoldePath, pageImage.PageNumber);

                    using (var stream = pageImage.Stream)
                    using (FileStream fileStream = new FileStream(pageImageName, FileMode.Create))
                    {
                        stream.Seek(0, SeekOrigin.Begin);
                        stream.CopyTo(fileStream);
                    }

                    var baseUrl = Request.Url.Scheme + "://" + Request.Url.Authority + Request.ApplicationPath.TrimEnd('/') + "/";
                    urls.Add(string.Format("{0}Content/TempStorage/{1}/{2}.png", baseUrl, request.Path, pageImage.PageNumber));
                }

                result.imageUrls = urls.ToArray();
            }
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

            var serializedData = serializer.Serialize(result);
            return Content(serializedData, "application/json");
        }

        public ActionResult LoadFileBrowserTreeData(LoadFileBrowserTreeDataParameters parameters)
        {
            var request = new LoadFileBrowserTreeRequest { Path = _storagePath };

            var tree = _htmlHandler.LoadFileBrowserTreeData(request);

            var data = new FileBrowserTreeDataResponse
            {
                nodes = ToFileTreeNodes(tree.Nodes).ToArray(),
                count = tree.Nodes.Count
            };

            JavaScriptSerializer serializer = new JavaScriptSerializer();

            var serializedData = serializer.Serialize(data);
            return Content(serializedData, "application/json");
        }

        public ActionResult GetImageUrls(GetImageUrlsParameters parameters)
        {
            if (string.IsNullOrEmpty(parameters.Path))
            {
                GetImageUrlsResponse empty = new GetImageUrlsResponse { imageUrls = new string[0] };

                var serialized = new JavaScriptSerializer().Serialize(empty);
                return Content(serialized, "application/json");
            }

            var imageOptions = new ImageOptions();
            var imagePages = _imageHandler.GetPages(new FileDescription { Guid = parameters.Path, Name = parameters.Path }, imageOptions);

            // Save images some where and provide urls
            var urls = new List<string>();
            var tempFolderPath = Path.Combine(Server.MapPath("~"), "Content", "TempStorage");

            foreach (var pageImage in imagePages)
            {
                var docFoldePath = Path.Combine(tempFolderPath, parameters.Path);

                if (!Directory.Exists(docFoldePath))
                    Directory.CreateDirectory(docFoldePath);

                var pageImageName = string.Format("{0}\\{1}.png", docFoldePath, pageImage.PageNumber);

                using (var stream = pageImage.Stream)
                using (FileStream fileStream = new FileStream(pageImageName, FileMode.Create))
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }

                var baseUrl = Request.Url.Scheme + "://" + Request.Url.Authority + Request.ApplicationPath.TrimEnd('/') + "/";
                urls.Add(string.Format("{0}Content/TempStorage/{1}/{2}.png", baseUrl, parameters.Path, pageImage.PageNumber));
            }

            GetImageUrlsResponse result = new GetImageUrlsResponse { imageUrls = urls.ToArray() };

            var serializedData = new JavaScriptSerializer().Serialize(result);
            return Content(serializedData, "application/json");
        }



        //public ActionResult GetFile(GetFileParameters parameters)
        //{
        //    var response = _htmlHandler.GetFile(parameters.Path);

        //    if (response == null)
        //        return new EmptyResult();

        //    return File(response.Stream, "application/octet-stream", parameters.DisplayName);
        //}

        //public ActionResult GetPdfWithPrintDialog(GetFileParameters parameters)
        //{
        //    throw new Exception();
        //}

        private Watermark GetWatermark(ViewDocumentParameters request)
        {
            if (string.IsNullOrWhiteSpace(request.WatermarkText))
                return null;

            return new Watermark(request.WatermarkText)
                {
                    Color = request.WatermarkColor.HasValue
                        ? Color.FromArgb(request.WatermarkColor.Value)
                        : Color.Red,
                    Position = ToWatermarkPosition(request.WatermarkPosition),
                    Width = request.WatermarkWidth
                };
        }

        private GroupDocs.Viewer.Domain.WatermarkPosition? ToWatermarkPosition(WatermarkPosition? watermarkPosition)
        {
            if (!watermarkPosition.HasValue)
                return GroupDocs.Viewer.Domain.WatermarkPosition.Diagonal;

            switch (watermarkPosition.Value)
            {
                case WatermarkPosition.Diagonal:
                    return GroupDocs.Viewer.Domain.WatermarkPosition.Diagonal;
                case WatermarkPosition.TopLeft:
                    return GroupDocs.Viewer.Domain.WatermarkPosition.TopLeft;
                case WatermarkPosition.TopCenter:
                    return GroupDocs.Viewer.Domain.WatermarkPosition.TopCenter;
                case WatermarkPosition.TopRight:
                    return GroupDocs.Viewer.Domain.WatermarkPosition.TopRight;
                case WatermarkPosition.BottomLeft:
                    return GroupDocs.Viewer.Domain.WatermarkPosition.BottomLeft;
                case WatermarkPosition.BottomCenter:
                    return GroupDocs.Viewer.Domain.WatermarkPosition.BottomCenter;
                case WatermarkPosition.BottomRight:
                    return GroupDocs.Viewer.Domain.WatermarkPosition.BottomRight;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private List<FileBrowserTreeNode> ToFileTreeNodes(IEnumerable<BrowserTreeNode> nodes)
        {
            return nodes.Select(_ =>
                new FileBrowserTreeNode
                {
                    path = _.Name,
                    docType = _.DocumentType,
                    fileType = _.FileType,
                    name = _.Name,
                    size = _.Size,
                    modifyTime = (long)(_.DateModified - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds,
                    type = _.Type,
                    nodes = ToFileTreeNodes(_.Nodes)
                })
                .ToList();
        }
    }
}