using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Web.Hosting;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace BrandingWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        public string themeName = "Branding";
        public string colorFilePath = System.IO.File.Exists(HostingEnvironment.MapPath("~/Theme/branding.spcolor")) ? HostingEnvironment.MapPath("~/Theme/branding.spcolor") : "";
        public string colorFileUrl;
        public string fontFilePath = System.IO.File.Exists(HostingEnvironment.MapPath("~/Theme/branding.spfont")) ? HostingEnvironment.MapPath("~/Theme/branding.spfont") : "";
        public string fontFileUrl;
        public string backgroundFilePath = System.IO.File.Exists(HostingEnvironment.MapPath("~/Theme/branding_bg.jpg")) ? HostingEnvironment.MapPath("~/Theme/branding_bg.jpg") : "";
        public string backgroundFileUrl;
        public string masterPageName = "seattle.master";
        public string masterPageFileUrl;
        

        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            switch (properties.EventType)
            {
                case SPRemoteEventType.AppInstalled:
                    HandleAppInstalled(properties);
                    break;
                case SPRemoteEventType.AppUninstalling:
                    HandleAppUninstalling(properties);
                    break;
                case SPRemoteEventType.WebProvisioned:
                    WebProvisionedEventReceiver(properties);
                    break;
            }

            return result;
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }

        private void HandleAppInstalled(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    Site site = clientContext.Site;
                    Web web = clientContext.Web;
                    clientContext.Load(site);
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    masterPageFileUrl = web.ServerRelativeUrl + "/_catalogs/masterpage/" + System.IO.Path.GetFileName(masterPageName);
                    colorFileUrl = string.IsNullOrEmpty(colorFilePath) ? "" : site.ServerRelativeUrl + "/_catalogs/theme/15/" + System.IO.Path.GetFileName(colorFilePath);
                    fontFileUrl = string.IsNullOrEmpty(fontFilePath) ? "" : site.ServerRelativeUrl + "/_catalogs/theme/15/" + System.IO.Path.GetFileName(fontFilePath);
                    backgroundFileUrl = string.IsNullOrEmpty(backgroundFilePath) ? "" : site.ServerRelativeUrl + "/_layouts/15/images/" + System.IO.Path.GetFileName(backgroundFilePath);

                    if (web.ServerRelativeUrl == site.ServerRelativeUrl)
                    {
                        UploadThemeFiles(web);
                    }

                    CreateCustomComposedLook(site, web);
                    ApplyTheme(web, colorFileUrl, fontFileUrl, backgroundFileUrl);
                    SetCurrentComposedLook(web, masterPageFileUrl, colorFileUrl, fontFileUrl, backgroundFileUrl);
                    
                    //AddWebProvisionedEventReceiverToHostWeb(web);
                }
            }
        }

        private void HandleAppUninstalling(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    Site site = clientContext.Site;
                    Web web = clientContext.Web;

                    if (site.RootWeb.ServerRelativeUrl == web.ServerRelativeUrl)
                    {
                        //delete theme files
                    }

                    ApplyTheme(web, site.ServerRelativeUrl + "/_catalogs/theme/15/palette001.spcolor", "", "");
                    SetCurrentComposedLook(web, web.ServerRelativeUrl + "/_catalogs/masterpage/seattle.master", site.ServerRelativeUrl + "/_catalogs/theme/15/palette001.spcolor", "", "");
                    DeleteCustomComposedLook(web);
                    //RemoveWebProvisionedEventReceiverFromHostWeb(web);
                }
            }
        }

        private void WebProvisionedEventReceiver(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    Site site = clientContext.Site;
                    Web web = clientContext.Web;

                    if (site.RootWeb.ServerRelativeUrl != web.ServerRelativeUrl)
                    {
                        try
                        {
                            Guid sideLoadingFeatureId = new Guid("AE3A1339-61F5-4f8f-81A7-ABD2DA956A7D");
                            //Add branding app to web
                            site.Features.Add(sideLoadingFeatureId, true, FeatureDefinitionScope.Site);

                            try
                            {
                                var appStream = System.IO.File.OpenRead(site.ServerRelativeUrl + "/add-ins/catalog/AppCatalog/Branding.app");
                                AppInstance app = web.LoadAndInstallApp(appStream);
                                clientContext.Load(app);
                                clientContext.ExecuteQuery();
                            }
                            catch
                            {
                                throw;
                            }

                            site.Features.Remove(sideLoadingFeatureId, true);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
        }

        private void UploadThemeFiles(Web web)
        {
            string[] files = new string[] { colorFilePath, fontFilePath, backgroundFilePath };
            List themesLibrary = web.GetCatalog(123);
            web.Context.Load(themesLibrary);
            web.Context.ExecuteQuery();
            Folder rootFolder = themesLibrary.RootFolder;
            web.Context.Load(rootFolder);
            web.Context.Load(rootFolder.Folders);
            web.Context.ExecuteQuery();
            foreach (Folder folder in rootFolder.Folders)
            {
                if (folder.Name == "15")
                {
                    foreach (string file in files)
                    {
                        if (!string.IsNullOrEmpty(file))
                        {
                            FileCreationInformation newFile = new FileCreationInformation();
                            newFile.Content = System.IO.File.ReadAllBytes(colorFilePath);
                            newFile.Url = folder.ServerRelativeUrl + "/" + System.IO.Path.GetFileName(colorFilePath);
                            newFile.Overwrite = true;
                            File uploadFile = folder.Files.Add(newFile);
                            web.Context.Load(uploadFile);
                            web.Context.ExecuteQuery();
                        }
                    }
                }
            }
        }

        private void CreateCustomComposedLook(Site site, Web web)
        {
            //Composed Looks List
            List list = web.GetCatalog(124);

            web.Context.Load(list);
            web.Context.ExecuteQuery();

            ListItemCollection themes = CheckThemeExists(web, list, themeName);

            if (themes.Count == 0)
            {
                ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                ListItem item = list.AddItem(itemInfo);

                item["Name"] = themeName;
                item["Title"] = themeName;
                item["ThemeUrl"] = colorFileUrl;
                item["FontSchemeUrl"] = fontFileUrl;
                item["ImageUrl"] = backgroundFileUrl;
                item["MasterPageUrl"] = masterPageFileUrl;
                item["DisplayOrder"] = 1;
                item.Update();
                web.Context.ExecuteQuery();
            }
        }

        private void SetCurrentComposedLook(Web web, string masterPageFileUrl, string colorFileUrl, string fontFileUrl, string backgroundFileUrl)
        {
            List list = web.GetCatalog(124);
            web.Context.Load(list);
            web.Context.ExecuteQuery();

            ListItemCollection themes = CheckThemeExists(web, list, "Current");

            if (themes.Count == 1)
            {
                ListItem item = themes.FirstOrDefault();
                FieldUrlValue masterPageUrl = new FieldUrlValue();
                FieldUrlValue themeUrl = new FieldUrlValue();
                FieldUrlValue fontSchemeUrl = new FieldUrlValue();
                FieldUrlValue imageUrl = new FieldUrlValue();

                masterPageUrl.Description = string.IsNullOrEmpty(masterPageFileUrl) ? "" : masterPageFileUrl;
                masterPageUrl.Url = string.IsNullOrEmpty(masterPageFileUrl) ? "" : masterPageFileUrl;
                themeUrl.Description = string.IsNullOrEmpty(colorFileUrl) ? "" : colorFileUrl;
                themeUrl.Url = string.IsNullOrEmpty(colorFileUrl) ? "" : colorFileUrl;
                fontSchemeUrl.Description = string.IsNullOrEmpty(fontFileUrl) ? "" : fontFileUrl;
                fontSchemeUrl.Url = string.IsNullOrEmpty(fontFileUrl) ? "" : fontFileUrl;
                imageUrl.Description = string.IsNullOrEmpty(backgroundFileUrl) ? "" : backgroundFileUrl;
                imageUrl.Url = string.IsNullOrEmpty(backgroundFileUrl) ? "" : backgroundFileUrl;

                item["MasterPageUrl"] = masterPageUrl;
                item["ThemeUrl"] = themeUrl;
                item["FontSchemeUrl"] = fontSchemeUrl;
                item["ImageUrl"] = imageUrl;
                item.Update();
                web.Context.ExecuteQuery();
            }
        }

        private void DeleteCustomComposedLook(Web web)
        {
            //Composed Looks List
            List list = web.GetCatalog(124);

            web.Context.Load(list);
            web.Context.ExecuteQuery();

            ListItemCollection themes = CheckThemeExists(web, list, themeName);

            if (themes.Count == 1)
            {
                //CamlQuery query = new CamlQuery();
                //string camlString = @"<View><Query><Where><Eq><FieldRef Name='Name' /><Value Type='Text'>{0}</Value></Eq></Where></Query></View>";
                //camlString = string.Format(camlString, themeName);
                //query.ViewXml = camlString;
                //var items = list.GetItems(query);
                ListItem theme = themes.FirstOrDefault();
                theme.DeleteObject();
                web.Context.ExecuteQuery();
            }
        }

        private ListItemCollection CheckThemeExists(Web web, List list, string name)
        {
            CamlQuery query = new CamlQuery();
            string camlString = @"<View><Query><Where><Eq><FieldRef Name='Name' /><Value Type='Text'>{0}</Value></Eq></Where></Query></View>";
            camlString = string.Format(camlString, name);
            query.ViewXml = camlString;
            ListItemCollection items = list.GetItems(query);

            web.Context.Load(items);
            web.Context.ExecuteQuery();

            return items;
        }

        private void ApplyTheme(Web web, string colorFileUrl, string fontFileUrl, string backgroundFileUrl)
        {
            web.ApplyTheme(string.IsNullOrEmpty(colorFileUrl) ? "" : colorFileUrl, string.IsNullOrEmpty(fontFileUrl) ? null : fontFileUrl, string.IsNullOrEmpty(backgroundFileUrl) ? null : backgroundFileUrl, false);
            web.Context.ExecuteQuery();
        }

        private void AddWebProvisionedEventReceiverToHostWeb(Web web)
        {
            EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();
            receiver.EventType = EventReceiverType.WebProvisioned;

            OperationContext op = OperationContext.Current;
            Message msg = op.RequestContext.RequestMessage;
            receiver.ReceiverUrl = msg.Headers.To.ToString();

            receiver.ReceiverName = themeName;
            receiver.Synchronization = EventReceiverSynchronization.Synchronous;
            web.EventReceivers.Add(receiver);
            web.Context.ExecuteQuery();
        }

        private void RemoveWebProvisionedEventReceiverFromHostWeb(Web web)
        {
            web.Context.Load(web, x => x.EventReceivers);
            web.Context.ExecuteQuery();

            var receiver = web.EventReceivers.Where(x => x.ReceiverName == themeName).FirstOrDefault();

            try
            {
                receiver.DeleteObject();
                web.Context.ExecuteQuery();
            }
            catch (Exception ex)
            {

            }
        }
    }
}
