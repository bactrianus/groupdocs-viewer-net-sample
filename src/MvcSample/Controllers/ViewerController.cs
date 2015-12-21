using GroupDocs.Viewer.Config;
using GroupDocs.Viewer.Converter.Option;
using GroupDocs.Viewer.Domain;
using GroupDocs.Viewer.Domain.Requests;
using GroupDocs.Viewer.Handler;
using GroupDocs.Viewer.Helper;
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
        private readonly string _storagePath = @"C:\storage";

        public ViewerController()
        {
            var config = new ViewerConfig
            {
                StoragePath = _storagePath
            };

            _htmlHandler = new ViewerHtmlHandler(config);
            _imageHandler = new ViewerImageHandler(config);

            _htmlHandler.SetLicense(@"D:\GroupDocs.Viewer.lic");
            _imageHandler.SetLicense(@"D:\GroupDocs.Viewer.lic");
        }

        // GET: /Viewer/
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ViewDocument(ViewDocumentParameters request)
        {
            var docInfo = _htmlHandler.GetDocumentInfo(new GetDocumentInfoRequest(request.Path));
            var docInfoAsJson = new FileDataJsonSerializer(docInfo.FileData, new FileDataOptions()).Serialize();

            var result = new ViewDocumentResponse
            {
                documentDescription = docInfoAsJson,
                pageCss = new string[] { },
                lic = true
            };

            if (request.UseHtmlBasedEngine)
            {
                var htmlOptions = new HtmlOptions
                {
                    Watermark = new Watermark("Watermark for html")
                    {
                        Color = Color.Blue,
                        Position = GroupDocs.Viewer.Domain.WatermarkPosition.TopCenter
                    }
                };

                var htmlPages = _htmlHandler.GetPages(new FileDescription { Guid = request.Path, Name = request.Path }, htmlOptions);
                result.pageHtml = htmlPages.Select(_ => _.HtmlContent).ToArray();
            }
            else
            {
                var imageOptions = new ImageOptions
                {
                    Watermark = GetWatermark(request)
                };

                var imagePages = _imageHandler.GetPages(new FileDescription { Guid = request.Path, Name = request.Path }, imageOptions);

                // Save images some where and provide urls
                var urls = new List<string>();
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

            var serializedData = new JavaScriptSerializer().Serialize(result);
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

            string serializedData = serializer.Serialize(data);
            return Content(serializedData, "application/json");
        }

        public ActionResult GetImageUrls(GetImageUrlsParameters parameters)
        {
            if (string.IsNullOrEmpty(parameters.Path))
            {
                GetImageUrlsResponse empty = new GetImageUrlsResponse();
                empty.imageUrls = new string[0];

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
            GetImageUrlsResponse result = new GetImageUrlsResponse();
            result.imageUrls = urls.ToArray();

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

        private List<FileBrowserTreeNode> ToFileTreeNodes(List<BrowserTreeNode> nodes)
        {
            var result = new List<FileBrowserTreeNode>();

            foreach (var _ in nodes)
            {
                var x = new FileBrowserTreeNode
                {
                    path = _.Name,
                    docType = _.DocumentType,
                    fileType = _.FileType,
                    name = _.Name,
                    size = _.Size,
                    modifyTime = (long)(_.DateModified - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds,
                    type = _.Type
                };
                x.nodes = ToFileTreeNodes(_.Nodes);
                result.Add(x);
            }
            return result;
        }
    }
}