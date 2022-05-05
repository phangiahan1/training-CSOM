using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using Microsoft.SharePoint.News.DataModel;

namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();
                    Console.WriteLine($"Site {ctx.Web.Title}");

                    //[1.1] Using CSOM create a list name "CSOM Test"
                    //await CreateList(ctx, "CSOM Test", "using CSOM create a list");

                    string termGroupName = "city";
                    string termSetName = "city-han";

                    //[1.2] Create term set "city-han" in dev tenant
                    //await CreateTermSet(ctx, termGroupName, termSetName);

                    //[1.3] Create 2 term "Ho Chi Minh" and "Stockholm" in termset "city-han"
                    //await CreateTerm(ctx, termGroupName, termSetName, "Ho Chi Minh");
                    //await CreateTerm(ctx, termGroupName, termSetName, "Stockholm");

                    //[1.4] Create site fields "about" type text and field "city" type taxonomy
                    string groupName = "CSOM city projects";
                    //await CreateSiteFieldTypeText(ctx, "about", "about", groupName);
                    //await CreateSiteFieldTypeTaxonomy(ctx, "city", "city", groupName);

                    //[1.5] Create site content type "CSOM Test content type"
                    //      => add this to "CSOM test"
                    //      add fields "about" and "city" to this.
                    string ContentTypeId = "0x0101009189AB5D3D2647B580F011DA2F356FB2";
                    string ContentTypeGroupName = "CSOM city projects Content Types";
                    //await CreateContentType(ctx, "CSOM Test content type", ContentTypeId, ContentTypeGroupName);

                    //await AddContentTypeToList(ctx, "CSOM Test content type", "CSOM Test");

                    //await AddFieldToContentType(ctx, "about", ContentTypeId);
                    //await AddFieldToContentType(ctx, "city", ContentTypeId);

                    //[1.6] In list "CSOM test" set "CSOM Test content type" as default content type
                    //await SetDefaultContentType(ctx, "CSOM Test content type", "CSOM Test");

                    //[1.7] Create 5 list items to list with some value  in field "about" and "city"
                    await CreateListItem(ctx, "CSOM Test", "phan gia han", "Ho Chi Minh");

                    //[1.8] Update site field "about" set default value for it to"about default" then create 2 new list items


                    //[1.9] Update site field "city" set default value for it to"Ho Chi Minh" then create 2 new list items

                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task GetFieldTermValue(ClientContext Ctx, string termId)
        {
            //load term by id
            TaxonomySession session = TaxonomySession.GetTaxonomySession(Ctx);
            Term taxonomyTerm = session.GetTerm(new Guid(termId));
            Ctx.Load(taxonomyTerm, t => t.Labels,
                                   t => t.Name,
                                   t => t.Id);
            await Ctx.ExecuteQueryAsync();
        }

        private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx, string fieldName)
        {
            //var field = ctx.Web.Fields.GetByTitle(fieldName);

            //ctx.Load(field);
            //await ctx.ExecuteQueryAsync();

            //var taxField = ctx.CastTo<TaxonomyField>(field);

            //taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
            //{
            //    WssId = -1, // alway let it -1
            //    Label = "city-han",
            //    TermGuid = "ac5437a6-038e-4da7-8959-c9044ab38ce6"
            //});
            //item.Update();
            //await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("Test Term Set");

            var terms = termSet.GetAllTerms();

            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomLinqAsync(ClientContext ctx)
        {
            var fieldsQuery = from f in ctx.Web.Fields
                              where f.InternalName == "Test" ||
                                    f.TypeAsString == "TaxonomyFieldTypeMulti" ||
                                    f.TypeAsString == "TaxonomyFieldType"
                              select f;

            var fields = ctx.LoadQuery(fieldsQuery);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("Documents");

            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
}

        private static void GetTaxonomyFieldInfo(ClientContext clientContext, out Guid termStoreId, out Guid termSetId)
        {
            termStoreId = Guid.Empty;
            termSetId = Guid.Empty;

            TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();
            TermSetCollection termSets = termStore.GetTermSetsByName("city-han", 1033);

            clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            clientContext.Load(termStore, ts => ts.Id);
            clientContext.ExecuteQuery();

            termStoreId = termStore.Id;
            termSetId = termSets.FirstOrDefault()!.Id;
        }
    
        private static async Task CreateList(ClientContext ctx, string listName, string description)
        {
            Console.WriteLine("Using CSOM create a list name: " + listName);

            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = listName;
            creationInfo.Description = description;
            creationInfo.TemplateType = (int)ListTemplateType.GenericList; //Custom list

            List newList = ctx.Web.Lists.Add(creationInfo);
            ctx.Load(newList);
            // Execute the query to the server.
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateTermSet(ClientContext ctx, string termGroupName, string termSetName)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            if (termStore != null)
            {
                TermGroup termGroup = termStore.CreateGroup(termGroupName, Guid.NewGuid());
                TermSet myTermSet = termGroup.CreateTermSet(termSetName, Guid.NewGuid(), 1033);
                await ctx.ExecuteQueryAsync();
            }
        }
        private static async Task CreateTerm(ClientContext ctx, string termGroupName, string termSetName, string termName)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName(termGroupName);
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName(termSetName);
            // Add term
            termSet.CreateTerm(termName, 1033, Guid.NewGuid());
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateSiteFieldTypeText(ClientContext ctx, string displayName,string name, string groupName)
        {
            Web rootWeb = ctx.Site.RootWeb;
            // Mind the AddFieldOptions.AddFieldInternalNameHint flag
            rootWeb.Fields.AddFieldAsXml($"<Field DisplayName='{displayName}' Name='{name}' Group='{groupName}' Type='Text'/>",
                false,
                AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateSiteFieldTypeTaxonomy(ClientContext ctx, string displayName, string name, string groupName)
        {
            Web rootWeb = ctx.Site.RootWeb;
            Field field =  rootWeb.Fields.AddFieldAsXml($"<Field DisplayName='{displayName}' Name='{name}' Group='{groupName}' Type='TaxonomyFieldType'/>",
               false,
               AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryAsync();

            Guid termStoreId = Guid.Empty;
            Guid termSetId = Guid.Empty;
            GetTaxonomyFieldInfo(ctx, out termStoreId, out termSetId);

            // Retrieve as Taxonomy Field
            TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.Update();

            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateContentType(ClientContext ctx, string contentTypeName, string contentTypeId, string contentTypeGroup)
        {
            Web rootWeb = ctx.Site.RootWeb;

            // create by ID
            rootWeb.ContentTypes.Add(new ContentTypeCreationInformation
            {
                Name = contentTypeName,
                Id = contentTypeId,
                Group = contentTypeGroup
            });

            await ctx.ExecuteQueryAsync();
        }
        private static async Task AddContentTypeToList(ClientContext ctx, string contentTypeName, string listName)
        {
            ContentTypeCollection contentTypeCollection = ctx.Site.RootWeb.ContentTypes;

            // Get Content Types from Current web

            ctx.Load(contentTypeCollection);
            await ctx.ExecuteQueryAsync();

            // Get the content type from content type collection. Give the content type name over here
            ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == contentTypeName select contentType).FirstOrDefault();

            // Add existing content type on target list. Give target list name over here.
            List CSOMtestList = ctx.Web.Lists.GetByTitle(listName);
            CSOMtestList.ContentTypes.AddExistingContentType(targetContentType);
            CSOMtestList.Update();
            ctx.Web.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task AddFieldToContentType(ClientContext ctx, string fieldName, string contentTypeId)
        {
            ////add fields "about" and "city"
            Web rootWeb = ctx.Site.RootWeb;

            Field field = rootWeb.Fields.GetByInternalNameOrTitle(fieldName);
            ContentType CSOMContentType = rootWeb.ContentTypes.GetById(contentTypeId);

            CSOMContentType.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = field
            });
            CSOMContentType.Update(true);
            await ctx.ExecuteQueryAsync();
        }
        private static async Task SetDefaultContentType(ClientContext ctx, string contentTypeName, string listName)
        {
            Console.WriteLine("Set content type: " + contentTypeName + " as default in list: " + listName);
            //get list 
            List list = ctx.Web.Lists.GetByTitle(listName);

            //get content type collection
            ContentTypeCollection currentCtOrder = list.ContentTypes;
            ctx.Load(currentCtOrder);
            ctx.ExecuteQuery();

            IList<ContentTypeId> reverceOrder = new List<ContentTypeId>();
            foreach (ContentType ct in currentCtOrder)
            {
                if (ct.Name.Equals(contentTypeName))
                {
                    reverceOrder.Add(ct.Id);
                }
            }
            list.RootFolder.UniqueContentTypeOrder = reverceOrder;
            list.RootFolder.Update();
            list.Update();
            await ctx.ExecuteQueryAsync();

        }
        private static async Task CreateListItem(ClientContext ctx, string listName, string about, string city)
        {
            List oList = ctx.Web.Lists.GetByTitle(listName);

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = oList.AddItem(itemCreateInfo);


            oListItem["about"] = about;            
            oListItem["city"] = city;
            //https://www.c-sharpcorner.com/uploadfile/anavijai/programmatically-set-value-to-the-taxonomy-field-in-sharepoint-2010/

            oListItem.Update();

            await ctx.ExecuteQueryAsync();
        }
    }
}
