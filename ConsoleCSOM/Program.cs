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
                    //await CreateListItem(ctx, "CSOM Test", "Pham Thi Bich Tram", "Ho Chi Minh");
                    //await CreateListItem(ctx, "CSOM Test", "Trinh Gia Dinh", "Stockholm");
                    //await CreateListItem(ctx, "CSOM Test", "Xa Thi Man", "");

                    //[1.8] Update site field "about" set default value for it to"about default" then create 2 new list items
                    //await UpdateDefaultValueSiteFieldTypeText(ctx, "CSOM Test", "about", "about default");
                    //await CreateListItem(ctx, "CSOM Test", null, "");
                    //await CreateListItem(ctx, "CSOM Test", "Not null", "");

                    //[1.9] Update site field "city" set default value for it to"Ho Chi Minh" then create 2 new list items
                    //await UpdateDefaultValueSiteFieldTypeTaxonomy(ctx, "CSOM Test", "city", "Ho Chi Minh");
                    //await CreateListItem(ctx, "CSOM Test", "Cau 1.9", null);
                    //await CreateListItem(ctx, "CSOM Test", null, null);

                    //[2.1] Write CAML query to get list items where field “about” is not “about default”
                    //////Eq Equals
                    //////Neq Not equal
                    //////Gt Greater than
                    //////Geq Greater than or equal
                    //////Lt  Lower than
                    //////Leq Lower than or equal too
                    //////IsNull Is null
                    //////BeginsWith Begins with
                    //////Contains Contains

                    //await CAMLQueryWithWhere(ctx, "CSOM Test", "about", "Text", "Neq", "about default");

                    //[2.2] Create List View by CSOM order item newest in top and only show list item where “city” field has value “Ho Chi Minh”,
                    //View Fields: Id, Name, City, About

                    //await CreateListViewWithOrderNewestAndWhereCityInHoChiMinh(ctx, "CSOM Test");


                    //[2.3] Write function update list items in batch, try to update 2 items every time and update field “about” which have value
                    //“about default” to “Update script”. (CAML)


                    //[2.4] Create new field “author” type people in list “CSOM Test” then migrate all list items to set user admin to field “CSOM Test Author”

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

        //Exercise 1: 
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
        public static void UpdateTaxonomyField(ClientContext ctx, List list, ListItem listItem, string fieldName, string fieldValue)
        {
            Console.WriteLine("UpdateTaxonomyField");
            Field field = list.Fields.GetByInternalNameOrTitle(fieldName);
            TaxonomyField txField = ctx.CastTo<TaxonomyField>(field);
            TaxonomyFieldValue termValue = new TaxonomyFieldValue();
            string[] term = fieldValue.Split('|');
            termValue.Label = term[0];
            termValue.TermGuid = term[1];
            //termValue.Label = "Ho Chi Minh";
            //termValue.TermGuid = "90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
            termValue.WssId = -1;
            txField.SetFieldValueByValue(listItem, termValue);
            listItem.Update();
            ctx.Load(listItem);
            ctx.ExecuteQuery();
        }
        private static async Task CreateListItem(ClientContext ctx, string listName, string about, string city)
        {
            List oList = ctx.Web.Lists.GetByTitle(listName);

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = oList.AddItem(itemCreateInfo);

            if(about != null)
            {
                oListItem["about"] = about;
            }
            
            string fieldValue;
            if(city == "Ho Chi Minh")
            {
                fieldValue = "Ho Chi Minh|90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
            } 
            else if(city == "Stockholm")
            {
                fieldValue = "Stockholm|f50c5a60-1411-447d-81ca-4242f11d5380";
            }
            else
            {
                // If user input wrong set default city is Ho Chi Minh
                fieldValue = "Ho Chi Minh|90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
            }
            UpdateTaxonomyField(ctx, oList, oListItem, "city", fieldValue);

            oListItem.Update();

            await ctx.ExecuteQueryAsync();
        }
        private static async Task UpdateDefaultValueSiteFieldTypeTextInList(ClientContext ctx, string listName, string fieldName, string defaultValueForField)
        {
            List list = ctx.Web.Lists.GetByTitle(listName);
            Field field = list.Fields.GetByInternalNameOrTitle(fieldName);
            field.DefaultValue = defaultValueForField;
            field.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task UpdateDefaultValueSiteFieldTypeText(ClientContext ctx, string listName, string fieldName, string defaultValueForField)
        {
            Web rootWeb = ctx.Site.RootWeb;
            //List list = ctx.Web.Lists.GetByTitle(listName);
            Field field = rootWeb.Fields.GetByInternalNameOrTitle(fieldName);
            field.DefaultValue = defaultValueForField;
            field.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task UpdateDefaultValueSiteFieldTypeTaxonomyInList(ClientContext ctx, string listName, string fieldName, string fieldValue)
        {
            List olist = ctx.Web.Lists.GetByTitle(listName);
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = olist.AddItem(itemCreateInfo);
            Field field = olist.Fields.GetByInternalNameOrTitle(fieldName);
            TaxonomyField txField = ctx.CastTo<TaxonomyField>(field);
            TaxonomyFieldValue termValue = new TaxonomyFieldValue();
            
            if(fieldValue == "Ho Chi Minh")
            {
                termValue.Label = "Ho Chi Minh";
                termValue.TermGuid = "90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
                termValue.WssId = -1;
            }
            else
            {
                termValue.Label = "Stockholm";
                termValue.TermGuid = "f50c5a60-1411-447d-81ca-4242f11d5380";
                termValue.WssId = -1;
            }
            var validatedValue = txField.GetValidatedString(termValue);

            await ctx.ExecuteQueryAsync();
            txField.DefaultValue = validatedValue.Value;
            txField.UserCreated = false;
            txField.UpdateAndPushChanges(true);
            oListItem.Update();
            ctx.Load(oListItem);
            await ctx.ExecuteQueryAsync();
        }
        private static async Task UpdateDefaultValueSiteFieldTypeTaxonomy(ClientContext ctx, string listName, string fieldName, string fieldValue)
        {
            Web rootWeb = ctx.Site.RootWeb;
            //List olist = ctx.Web.Lists.GetByTitle(listName);
            //ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            //ListItem oListItem = olist.AddItem(itemCreateInfo);
            Field field = rootWeb.Fields.GetByInternalNameOrTitle(fieldName);
            TaxonomyField txField = ctx.CastTo<TaxonomyField>(field);
            TaxonomyFieldValue termValue = new TaxonomyFieldValue();

            if (fieldValue == "Ho Chi Minh")
            {
                termValue.Label = "Ho Chi Minh";
                termValue.TermGuid = "90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
                termValue.WssId = -1;
            }
            else
            {
                termValue.Label = "Stockholm";
                termValue.TermGuid = "f50c5a60-1411-447d-81ca-4242f11d5380";
                termValue.WssId = -1;
            }
            var validatedValue = txField.GetValidatedString(termValue);

            await ctx.ExecuteQueryAsync();
            txField.DefaultValue = validatedValue.Value;
            txField.UserCreated = false;
            txField.UpdateAndPushChanges(true);
            rootWeb.Update();
            ctx.Load(rootWeb);
            await ctx.ExecuteQueryAsync();
        }
        
        //Exercise 2:
        private static async Task CAMLQueryWithWhere(ClientContext ctx, string listName, string fieldName, string fielsType, string Operators, string OperatorsValue)
        {
            List list = ctx.Web.Lists.GetByTitle(listName);

            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View>"
               + "<Query>"
               + $"<Where><{Operators}><FieldRef Name='{fieldName}' /><Value Type='{fielsType}'>{OperatorsValue}</Value></{Operators}></Where>"
               + "</Query>"
               + "</View>";
            // execute the query
            ListItemCollection listItems = list.GetItems(query);
            //foreach (ListItem item in listItems)
            //{
            //    Console.WriteLine(item);
            //}
            ctx.Load(listItems);
            await ctx.ExecuteQueryAsync();

            foreach (ListItem oListItem in listItems)
            {
                TaxonomyFieldValue taxFieldValue = oListItem["city"] as TaxonomyFieldValue;
                Console.WriteLine("about: {0}  - city: {1}",  oListItem["about"], taxFieldValue.Label);
            }


        }
        private static async Task CreateListViewWithOrderNewestAndWhereCityInHoChiMinh(ClientContext ctx, string listName)
        {
            List targetList = ctx.Web.Lists.GetByTitle(listName);

            ViewCollection viewCollection = targetList.Views;
            ctx.Load(viewCollection);

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = "View With Order And City Is HCM";

            // Specify type of the view. Below are the options

            // 1. none - The type of the list view is not specified
            // 2. html - Sspecifies an HTML list view type
            // 3. grid - Specifies a datasheet list view type
            // 4. calendar- Specifies a calendar list view type
            // 5. recurrence - Specifies a list view type that displays recurring events
            // 6. chart - Specifies a chart list view type
            // 7. gantt - Specifies a Gantt chart list view type

            viewCreationInformation.ViewTypeKind = ViewType.Grid;

            // You can optionally specify row limit for the view
            //viewCreationInformation.RowLimit = 10;

            // You can optionally specify a query as mentioned below.
            // Create one CAML query to filter list view and mention that query below
            viewCreationInformation.Query = "<Where><Eq><FieldRef Name = 'city' /><Value Type = 'Taxonomy'>Ho Chi Minh</Value></Eq></Where>";
            viewCreationInformation.Query = "<OrderBy><FieldRef Name='Created' Ascending='FALSE'/></OrderBy>";

            // Add all the fields over here with comma separated value as mentioned below
            // You can mention display name or internal name of the column
            string CommaSeparateColumnNames = "about,city";
            viewCreationInformation.ViewFields = CommaSeparateColumnNames.Split(',');

            View listView = viewCollection.Add(viewCreationInformation);
            ctx.ExecuteQuery();

            // Code to update the display name for the view.
            listView.Title = "View With Order And City Is HCM";

            // You can optionally specify Aggregation: Field references for totals columns or calculated columns
            //listView.Aggregations = "<FieldRef Name='Title' Type='COUNT'/>";

            listView.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CAMLQueryUpdateMutiListItems(ClientContext ctx, string listName, string fieldName, string currentValue, string newValue)
        {

        }
    }
}
