using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Utilities;
using Microsoft.SharePoint.Client.Publishing;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPublishingPages : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Publishing Pages"; }
        }

        static string pageLibrary = "Pages";
        static string pageLayout = "MyLayout";
        static string pageName = "Test.aspx";
        static string pageDisplayName = pageName.Split('.')[0];
        static string pageURL = string.Empty;


        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                PublishingWeb webPub = PublishingWeb.GetPublishingWeb(context, web);
                context.Load(webPub);
                //context.ExecuteQuery();
                
                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.RootFolder.WelcomePage, w => w.AllProperties);

                if (!web.IsPublishingWeb())
                    throw new Exception(String.Format("Site '{0}' is not a publishing site", web.ServerRelativeUrl));

                // Get Publishing Page Layouts
                List publishingLayouts = context.Site.RootWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                ListItemCollection allPageLayouts = publishingLayouts.GetItems(CamlQuery.CreateAllItemsQuery());
                context.Load(allPageLayouts, items => items.Include(item => item.DisplayName, item => item["Title"]));
                context.ExecuteQuery();

                if (webPub != null)
                {
                    foreach (var page in template.PublishingPages)
                    {
                        var name = parser.ParseString(page.Name);

                        if (!name.ToLower().EndsWith(".aspx"))
                        {
                            name = String.Format("{0}{1}", name, ".aspx");
                        }

                        page.Name = name;

                        var exists = true;
                        string urlFile = string.Empty;
                        Microsoft.SharePoint.Client.File file = null;
                        try
                        {
                            urlFile = UrlUtility.Combine(web.ServerRelativeUrl, new[] { "Pages", page.Name });
                            file = web.GetFileByServerRelativeUrl(urlFile);
                            web.Context.Load(file);
                            web.Context.ExecuteQuery();
                        }
                        catch (ServerException ex)
                        {
                            if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                            {
                                exists = false;
                            }
                        }

                        if (exists)
                        {
                            if (page.Overwrite)
                            {
                                //delete page
                                if (web.RootFolder.WelcomePage.Contains(page.Name))
                                    web.SetHomePage(string.Empty);

                                file.DeleteObject();
                                web.Context.ExecuteQueryRetry();

                                try
                                {
                                    AddPublishingPage(context, webPub, allPageLayouts, page);
                                }
                                catch (Exception ex)
                                {
                                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0__failed___1_____2_, name, ex.Message, ex.StackTrace);
                                }
                            }
                        }
                        else
                        {
                            //Create Page
                            try
                            {
                                AddPublishingPage(context, webPub, allPageLayouts, page);
                            }
                            catch (Exception ex)
                            {
                                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0__failed___1_____2_, name, ex.Message, ex.StackTrace);
                            }
                        }

                        if (page.WebParts != null & page.WebParts.Any())
                        {
                            var existingWebParts = web.GetWebParts(urlFile);

                            foreach (var webpart in page.WebParts)
                            {
                                if (existingWebParts.FirstOrDefault(w => w.WebPart.Title == webpart.Title) == null)
                                {
                                    WebPartEntity wpEntity = new WebPartEntity();
                                    wpEntity.WebPartTitle = webpart.Title;
                                    wpEntity.WebPartXml = parser.ParseString(webpart.Contents.Trim(new[] { '\n', ' ' }));
                                    wpEntity.WebPartZone = webpart.Zone;
                                    wpEntity.WebPartIndex = (int)webpart.Order;
                                    web.AddWebPartToWebPartPage(urlFile, wpEntity);
                                }
                            }
                            var allWebParts = web.GetWebParts(urlFile);
                            foreach (var webpart in allWebParts)
                            {
                                parser.AddToken(new WebPartIdToken(web, webpart.WebPart.Title, webpart.Id));
                            }
                        }


                        CheckInAndPublishPage(context, webPub, urlFile);

                    }
                }
            }
            return parser;
        }

        private void CheckInAndPublishPage(ClientContext context, PublishingWeb webPub, string fileUrl)
        {
            //get the home page
            Microsoft.SharePoint.Client.File home = context.Web.GetFileByServerRelativeUrl(fileUrl);
            home.CheckIn(string.Empty, CheckinType.MajorCheckIn);
            home.Publish(string.Empty);
        }

        private static void AddPublishingPage(ClientContext context, PublishingWeb webPub, ListItemCollection allPageLayouts, OfficeDevPnP.Core.Framework.Provisioning.Model.PublishingPage page)
        {

            ListItem layout = null;

            //try to get layout by Title
            layout = allPageLayouts.Where(x => x["Title"] != null && x["Title"].ToString().Equals(page.Layout)).FirstOrDefault();

            //try to get layout by DisplayName
            if (layout == null)
                layout = allPageLayouts.Where(x => x.DisplayName == page.Layout).FirstOrDefault();

            //we need to have a layout for a publishing page
            if (layout == null)
                throw new ArgumentNullException(string.Format("Layout '{0}' for page {1} can not be found.", page.Layout, page.Name));

            context.Load(layout);

            // Create a publishing page    
            PublishingPageInformation publishingPageInfo = new PublishingPageInformation();
            publishingPageInfo.Name = page.Name;
            publishingPageInfo.PageLayoutListItem = layout;

            Microsoft.SharePoint.Client.Publishing.PublishingPage publishingPage = webPub.AddPublishingPage(publishingPageInfo);

            //publishingPage.ListItem.File.Approve(string.Empty);

            context.Load(publishingPage);
            context.Load(publishingPage.ListItem.File, obj => obj.ServerRelativeUrl);
            context.ExecuteQuery();
        }


        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Impossible to return all files in the site currently

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {

            return template;
        }


        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.PublishingPages.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }
    }
}
