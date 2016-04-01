using GroupDocs.Viewer.Config;
using GroupDocs.Viewer.Converter.Options;
using GroupDocs.Viewer.Domain;
using GroupDocs.Viewer.Domain.Html;
using GroupDocs.Viewer.Domain.Options;
using GroupDocs.Viewer.Handler;
using MvcSample.Helpers;
using MvcSample.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using WatermarkPosition = MvcSample.Models.WatermarkPosition;

namespace MvcSample.Controllers
{
    public class ViewerController : Controller
    {
        private readonly ViewerHtmlHandler _htmlHandler;
        private readonly ViewerImageHandler _imageHandler;

        // App_Data folder path
        private readonly string _storagePath = AppDomain.CurrentDomain.GetData("DataDirectory").ToString();
        private readonly string _tempPath = AppDomain.CurrentDomain.GetData("DataDirectory") + "\\Temp";

        public ViewerController()
        {
            var config = new ViewerConfig
            {
                StoragePath = _storagePath,
                TempPath = _tempPath,
                UseCache = false
            };

            _htmlHandler = new ViewerHtmlHandler(config);
            _imageHandler = new ViewerImageHandler(config);
        }

        // GET: /Viewer/
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ViewDocument(ViewDocumentParameters request)
        {
            if (Utils.IsValidUrl(request.Path))
                PrepareUrl(request);

            var fileName = Path.GetFileName(request.Path);

            var result = new ViewDocumentResponse
            {
                pageCss = new string[] { },
                lic = true,
                pdfDownloadUrl = GetPdfDownloadUrl(request),
                pdfPrintUrl = GetPdfPrintUrl(request),
                url = GetFileUrl(request),
                path = request.Path,
                name = fileName
            };

            if (request.UseHtmlBasedEngine)
                ViewDocumentAsHtml(request, result, fileName);
            else
                ViewDocumentAsImage(request, result, fileName);

            return ToJsonResult(result);
        }

        public ActionResult LoadFileBrowserTreeData(LoadFileBrowserTreeDataParameters parameters)
        {
            var path = _storagePath;
            if (!string.IsNullOrEmpty(parameters.Path))
                path = Path.Combine(path, parameters.Path);

            var request = new FileTreeOptions(path);
            var tree = _htmlHandler.LoadFileTree(request);

            var result = new FileBrowserTreeDataResponse
            {
                nodes = Utils.ToFileTreeNodes(parameters.Path, tree.FileTree).ToArray(),
                count = tree.FileTree.Count
            };

            return ToJsonResult(result);
        }

        public ActionResult GetImageUrls(GetImageUrlsParameters parameters)
        {
            if (string.IsNullOrEmpty(parameters.Path))
            {
                var empty = new GetImageUrlsResponse { imageUrls = new string[0] };

                var serialized = new JavaScriptSerializer().Serialize(empty);
                return Content(serialized, "application/json");
            }

            var imageOptions = new ImageOptions();
            var imagePages = _imageHandler.GetPages(parameters.Path, imageOptions);

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
                using (var fileStream = new FileStream(pageImageName, FileMode.Create))
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }

                var baseUrl = Request.Url.Scheme + "://" + Request.Url.Authority + Request.ApplicationPath.TrimEnd('/') +
                              "/";
                urls.Add(string.Format("{0}Content/TempStorage/{1}/{2}.png", baseUrl, parameters.Path,
                    pageImage.PageNumber));
            }

            var result = new GetImageUrlsResponse { imageUrls = urls.ToArray() };
            return ToJsonResult(result);
        }

        public ActionResult GetFile(GetFileParameters parameters)
        {
            var displayName = string.IsNullOrEmpty(parameters.DisplayName)
                ? Path.GetFileName(parameters.Path)
                : Uri.EscapeDataString(parameters.DisplayName);

            Stream fileStream;
            if (parameters.GetPdf)
            {
                displayName = Path.ChangeExtension(displayName, "pdf");

                var options = new PdfFileOptions
                {
                    Guid = parameters.Path,
                    Watermark = Utils.GetWatermark(parameters.WatermarkText, parameters.WatermarkColor, parameters.WatermarkPosition, parameters.WatermarkWidth),
                };

                if (parameters.IsPrintable)
                    options.AddPrintAction = true;

                if (parameters.SupportPageRotation)
                    options.Transformations |= Transformation.Rotate;

                options.Transformations |= Transformation.Reorder;

                var pdfFileResponse = _htmlHandler.GetPdfFile(options);
                fileStream = pdfFileResponse.Stream;
            }
            else
            {
                var fileResponse = _htmlHandler.GetFile(parameters.Path);
                fileStream = fileResponse.Stream;
            }

            //jquery.fileDownload uses this cookie to determine that a file download has completed successfully
            Response.SetCookie(new HttpCookie("jqueryFileDownloadJSForGD", "true") { Path = "/" });

            fileStream.Position = 0;
            using (var ms = new MemoryStream())
            {
                fileStream.CopyTo(ms);
                return File(ms.ToArray(), "application/octet-stream", displayName);
            }
        }

        public ActionResult GetPdfWithPrintDialog(GetFileParameters parameters)
        {
            var displayName = string.IsNullOrEmpty(parameters.DisplayName)
                ? Path.GetFileName(parameters.Path)
                : Uri.EscapeDataString(parameters.DisplayName);

            var options = new PdfFileOptions
            {
                Guid = parameters.Path,
                Watermark = Utils.GetWatermark(parameters.WatermarkText, parameters.WatermarkColor, parameters.WatermarkPosition, parameters.WatermarkWidth)
            };

            if (parameters.IsPrintable)
                options.AddPrintAction = true;

            if (parameters.SupportPageRotation)
                options.Transformations |= Transformation.Rotate;

            options.Transformations |= Transformation.Reorder;


            var response = _htmlHandler.GetPdfFile(options);

            var contentDispositionString = new ContentDisposition { FileName = displayName, Inline = true }.ToString();
            Response.AddHeader("Content-Disposition", contentDispositionString);

            return File(((MemoryStream)response.Stream).ToArray(), "application/pdf");
        }

        public string GetFileUrl(string path, bool getPdf, bool isPrintable, string fileDisplayName = null,
          string watermarkText = null, int? watermarkColor = null,
          WatermarkPosition? watermarkPosition = WatermarkPosition.Diagonal, float? watermarkWidth = 0,
          bool ignoreDocumentAbsence = false,
          bool useHtmlBasedEngine = false,
          bool supportPageRotation = false)
        {
            var queryString = HttpUtility.ParseQueryString(string.Empty);
            queryString["path"] = path;
            if (!isPrintable)
            {
                queryString["getPdf"] = getPdf.ToString().ToLower();

                if (fileDisplayName != null)
                    queryString["displayName"] = fileDisplayName;
            }

            if (watermarkText != null)
            {
                queryString["watermarkText"] = watermarkText;
                queryString["watermarkColor"] = watermarkColor.ToString();

                if (watermarkPosition.HasValue)
                    queryString["watermarkPosition"] = watermarkPosition.ToString();

                if (watermarkWidth.HasValue)
                    queryString["watermarkWidth"] = ((float)watermarkWidth).ToString(CultureInfo.InvariantCulture);
            }

            if (ignoreDocumentAbsence)
            {
                queryString["ignoreDocumentAbsence"] = ignoreDocumentAbsence.ToString().ToLower();
            }

            queryString["useHtmlBasedEngine"] = useHtmlBasedEngine.ToString().ToLower();
            queryString["supportPageRotation"] = supportPageRotation.ToString().ToLower();

            var handlerName = isPrintable ? "GetPdfWithPrintDialog" : "GetFile";

            var baseUrl = Request.Url.Scheme + "://" + Request.Url.Authority + Request.ApplicationPath.TrimEnd('/') + "/document-viewer/";

            var fileUrl = string.Format("{0}{1}?{2}", baseUrl, handlerName, queryString);
            return fileUrl;
        }

        public ActionResult GetDocumentPageHtml(GetDocumentPageHtmlParameters parameters)
        {
            if (Utils.IsValidUrl(parameters.Path))
                parameters.Path = Utils.GetFilenameFromString(parameters.Path);

            if (String.IsNullOrWhiteSpace(parameters.Path))
                throw new ArgumentException("A document path must be specified", "path");

            List<string> cssList;

            var htmlOptions = new HtmlOptions
            {
                PageNumber = parameters.PageIndex + 1,
                CountPagesToConvert = 1,
                IsResourcesEmbedded = false,
                HtmlResourcePrefix = string.Format(
                    "/document-viewer/GetResourceForHtml?documentPath={0}", parameters.Path) +
                                     "&pageNumber={page-number}&resourceName=",
            };

            var htmlPages = GetHtmlPages(parameters.Path, htmlOptions, out cssList);

            var pageHtml = htmlPages.Count > 0 ? htmlPages[0].HtmlContent : null;
            var pageCss = cssList.Count > 0 ? cssList[0] : null;

            var result = new { pageHtml, pageCss };
            return ToJsonResult(result);
        }

        public ActionResult GetResourceForHtml(GetResourceForHtmlParameters parameters)
        {
            var resource = new HtmlResource
            {
                ResourceName = parameters.ResourceName,
                ResourceType = Utils.GetResourceType(parameters.ResourceName),
                DocumentPageNumber = parameters.PageNumber
            };
            var stream = _htmlHandler.GetResource(parameters.DocumentPath, resource);

            if (stream == null || stream.Length == 0)
                return new HttpStatusCodeResult((int)HttpStatusCode.Gone);

            return File(stream, Utils.GetImageMimeTypeFromFilename(parameters.ResourceName));
        }


        private List<PageHtml> GetHtmlPages(string filePath, HtmlOptions htmlOptions, out List<string> cssList)
        {
            var htmlPages = _htmlHandler.GetPages(filePath, htmlOptions);

            cssList = new List<string>();
            foreach (var page in htmlPages)
            {
                var indexOfBodyOpenTag = page.HtmlContent.IndexOf("<body>", StringComparison.InvariantCultureIgnoreCase);

                if (indexOfBodyOpenTag > 0)
                    page.HtmlContent = page.HtmlContent.Substring(indexOfBodyOpenTag + "<body>".Length);

                var indexOfBodyCloseTag = page.HtmlContent.IndexOf("</body>", StringComparison.InvariantCultureIgnoreCase);

                if (indexOfBodyCloseTag > 0)
                    page.HtmlContent = page.HtmlContent.Substring(0, indexOfBodyCloseTag);

                foreach (var resource in page.HtmlResources.Where(_ => _.ResourceType == HtmlResourceType.Style))
                {
                    var cssStream = _htmlHandler.GetResource(filePath, resource);
                    var text = new StreamReader(cssStream).ReadToEnd();
                    text = text.Replace("url(\"",
                        string.Format("url(\"/document-viewer/GetResourceForHtml?documentPath={0}&pageNumber={1}&resourceName=",
                        filePath, page.PageNumber));
                    cssList.Add(text);
                }
            }
            return htmlPages;
        }

        private ActionResult ToJsonResult(object result)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var serializedData = serializer.Serialize(result);
            return Content(serializedData, "application/json");
        }

        private void PrepareUrl(ViewDocumentParameters request)
        {
            var fileNameFromUrl = Utils.GetFilenameFromString(request.Path);
            var filePath = Path.Combine(_storagePath, fileNameFromUrl);

            using (new InterProcessLock(filePath))
                Utils.DownloadFile(request.Path, filePath);

            request.Path = fileNameFromUrl;
        }

        private void ViewDocumentAsImage(ViewDocumentParameters request, ViewDocumentResponse result, string fileName)
        {
            var docInfo = _imageHandler.GetDocumentInfo(new DocumentInfoOptions(request.Path));
            var fileData = new FileData
            {
                DateCreated = DateTime.Now,
                DateModified = docInfo.LastModificationDate,
                PageCount = docInfo.Pages.Count,
                Pages = new List<PageData>(),
                MaxHeight = 900,
                MaxWidth = 600
            };

            for (var i = 0; i < docInfo.Pages.Count; i++)
            {
                var page = new PageData
                {
                    Angle = 0,
                    Height = 900,
                    Number = i,
                    //Name = "page" + i,
                    Rows = new List<RowData>(),
                    Width = 600
                };

                fileData.Pages.Add(page);
            }

            result.documentDescription = new FileDataJsonSerializer(fileData, new FileDataOptions()).Serialize(true);
            result.docType = docInfo.DocumentType;
            result.fileType = docInfo.FileType;

            var imageOptions = new ImageOptions
            {
                Watermark = Utils.GetWatermark(request.WatermarkText, request.WatermarkColor, request.WatermarkPosition, request.WatermarkWidth)
            };
            var imagePages = _imageHandler.GetPages(fileName, imageOptions);

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
                using (var fileStream = new FileStream(pageImageName, FileMode.Create))
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }

                var baseUrl = Request.Url.Scheme + "://" + Request.Url.Authority +
                              Request.ApplicationPath.TrimEnd('/') + "/";
                urls.Add(string.Format("{0}Content/TempStorage/{1}/{2}.png", baseUrl, request.Path,
                    pageImage.PageNumber));
            }

            result.imageUrls = urls.ToArray();
        }

        private void ViewDocumentAsHtml(ViewDocumentParameters request, ViewDocumentResponse result, string fileName)
        {
            var docInfo = _htmlHandler.GetDocumentInfo(new DocumentInfoOptions(request.Path));
            var fileData = new FileData
            {
                DateCreated = DateTime.Now,
                DateModified = docInfo.LastModificationDate,
                PageCount = docInfo.Pages.Count,
                Pages = new List<PageData>(),
                MaxHeight = 900,
                MaxWidth = 600
            };

            for (var i = 0; i < docInfo.Pages.Count; i++)
            {
                var page = new PageData
                {
                    Angle = 0,
                    Height = 900,
                    Number = i,
                    //Name = "page" + i,
                    Rows = new List<RowData>(),
                    Width = 600
                };

                fileData.Pages.Add(page);
            }

            result.documentDescription = new FileDataJsonSerializer(fileData, new FileDataOptions()).Serialize(false);
            result.docType = docInfo.DocumentType;
            result.fileType = docInfo.FileType;

            var htmlOptions = new HtmlOptions
            {
                IsResourcesEmbedded = false,
                HtmlResourcePrefix = string.Format(
                "/document-viewer/GetResourceForHtml?documentPath={0}", fileName) + "&pageNumber={page-number}&resourceName=",
                Watermark = Utils.GetWatermark(request.WatermarkText, request.WatermarkColor, request.WatermarkPosition, request.WatermarkWidth)
            };

            if (request.PreloadPagesCount.HasValue && request.PreloadPagesCount.Value > 0)
            {
                htmlOptions.PageNumber = 1;
                htmlOptions.CountPagesToConvert = request.PreloadPagesCount.Value;
            }

            List<string> cssList;
            var htmlPages = GetHtmlPages(fileName, htmlOptions, out cssList);
            result.pageHtml = htmlPages.Select(_ => _.HtmlContent).ToArray();
            result.pageCss = cssList.ToArray();

            //NOTE: Fix for incomplete cells document
            for (var i = 0; i < result.pageHtml.Length; i++)
            {
                var html = result.pageHtml[i];
                var indexOfScript = html.IndexOf("script", StringComparison.InvariantCultureIgnoreCase);
                if (indexOfScript > 0)
                    result.pageHtml[i] = html.Substring(0, indexOfScript);
            }
        }

        private string GetFileUrl(ViewDocumentParameters request)
        {
            return GetFileUrl(request.Path, false, false, request.FileDisplayName);
        }

        private string GetPdfPrintUrl(ViewDocumentParameters request)
        {
            return GetFileUrl(request.Path, true, true, request.FileDisplayName,
                request.WatermarkText, request.WatermarkColor,
                request.WatermarkPosition, request.WatermarkWidth,
                request.IgnoreDocumentAbsence,
                request.UseHtmlBasedEngine, request.SupportPageRotation);
        }

        private string GetPdfDownloadUrl(ViewDocumentParameters request)
        {
            return GetFileUrl(request.Path, true, false, request.FileDisplayName,
                request.WatermarkText, request.WatermarkColor,
                request.WatermarkPosition, request.WatermarkWidth,
                request.IgnoreDocumentAbsence,
                request.UseHtmlBasedEngine, request.SupportPageRotation);
        }
    }
}