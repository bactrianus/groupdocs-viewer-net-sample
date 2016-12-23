# GroupDocs.Viewer for .NET ASP.NET MVC Sample

[GroupDocs.Viewer for .NET](https://www.nuget.org/packages/groupdocs-viewer-dotnet/) ASP.NET MVC Sample contains sample application 
that demonstrates how GroupDocs.Viewer for .NET can be used in web application to render different file formats into image or html.

## How to use

Clone or download the application, restore NuGet packages and run the application.

## How to set license

To set license update _licensePath in Global.asax.cs:

```csharp
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace MvcSample
{
    public class MvcApplication : System.Web.HttpApplication
    {
        private readonly string _licensePath = "c:\\licenses\\GroupDocs.Viewer.lic";

        protected void Application_Start()
        {
            //...

            //Set GroupDocs.Viewer License
            GroupDocs.Viewer.License license = new GroupDocs.Viewer.License();
            license.SetLicense(_licensePath);
        }
    }
}
```

## Resources

+ **[GroupDocs.Viewer for .NET Home](http://groupdocs.com/dot-net/document-viewer-library)**
+ **[GroupDocs.Viewer for .NET API](http://groupdocs.com/api/net/viewer)**
+ **[GroupDocs.Viewer for .NET Documentation](http://groupdocs.com/docs/display/viewernet/Introducing+GroupDocs.Viewer+for+.NET)**

