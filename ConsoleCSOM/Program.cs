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
using System.IO;
using ListItem = Microsoft.SharePoint.Client.ListItem;
using View = Microsoft.SharePoint.Client.View;
using System.Text;
using Microsoft.SharePoint.Client.UserProfiles;

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
                    string LIST_NAME = "CSOM Test";
                    ClientContext ctx = GetContext(clientContextHelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();
                    Console.WriteLine($"Site {ctx.Web.Title}");

                    //[1.1] Using CSOM create a list name "CSOM Test"
                    //await CreateList(ctx, LIST_NAME, "using CSOM create a list");

                    string termGroupName = "city";
                    string termSetName = "city-han";
                    string HO_CHI_MINH = "Ho Chi Minh";
                    string STOCKHOLM = "Stockholm";

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
                    //string ContentTypeName = "CSOM Test content type";
                    string ContentTypeId = "0x0101009189AB5D3D2647B580F011DA2F356FB2";
                    //string ContentTypeId = "0x0101006754AB5D3D2647B580F011DA2F356FB9";
                    string ContentTypeGroupName = "CSOM city projects Content Types";
                    string ContentTypeName = "CSOM Test content type";

                    //await CreateContentType(ctx, ContentTypeName, ContentTypeId, ContentTypeGroupName);
                    //await AddContentTypeToList(ctx, ContentTypeName, LIST_NAME);

                    //await AddFieldToContentType(ctx, "about", ContentTypeId);
                    //await AddFieldToContentType(ctx, "city", ContentTypeId);

                    //[1.6] In list "CSOM test" set "CSOM Test content type" as default content type
                    //await SetDefaultContentType(ctx, ContentTypeName, LIST_NAME);

                    //[1.7] Create 5 list items to list with some value  in field "about" and "city"
                    //await CreateListItem(ctx, LIST_NAME, "Duong Duy Tan", HO_CHI_MINH);
                    //await CreateListItem(ctx, LIST_NAME, "Duong Tran Khanh Van", STOCKHOLM);
                    //await CreateListItem(ctx, LIST_NAME, "Phan Tu Hai", null);
                    //await CreateListItem(ctx, LIST_NAME, "Vo Thi Ngoc Mai", HO_CHI_MINH);
                    //await CreateListItem(ctx, LIST_NAME, "Phan Gia Hân", STOCKHOLM);

                    //[1.8] Update site field "about" set default value for it to"about default" then create 2 new list items
                    //await UpdateDefaultValueSiteFieldTypeTextInList(ctx, LIST_NAME, "about", "about default 111");
                    //await CreateListItem(ctx, LIST_NAME, null, "");
                    //await CreateListItem(ctx, LIST_NAME, "Not null", "");

                    //[1.9] Update site field "city" set default value for it to"Ho Chi Minh" then create 2 new list items
                    //await UpdateDefaultValueSiteFieldTypeTaxonomy(ctx, LIST_NAME, "city", HO_CHI_MINH);
                    //await CreateListItem(ctx, LIST_NAME, "Cau 1.9", null);
                    //await CreateListItem(ctx, LIST_NAME, null, null);

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

                    //await CAMLQueryWithWhere(ctx, LIST_NAME, "about", "Text", "Eq", "about default 111");

                    //[2.2] Create List View by CSOM order item newest in top and only show list item where “city” field has value “Ho Chi Minh”,
                    //View Fields: Id, Name, City, About

                    string NameViewWithOrderNewestAndWhereCityInHoChiMinh = "View With Order And City Is HCM";
                    //await CreateListViewWithOrderNewestAndWhereCityInHoChiMinh(ctx, LIST_NAME, NameViewWithOrderNewestAndWhereCityInHoChiMinh);

                    //[2.3] Write function update list items in batch, try to update 2 items every time and update field “about” which have value
                    //“about default” to “Update script”. (CAML)

                    //await CAMLQueryUpdateMutiListItems(ctx, LIST_NAME, "about", "Text", "Update script", "about default");
                    //await CAMLQueryUpdateMutiListItems(ctx, LIST_NAME, "about", "Text", "about default 111", "Update script");

                    //[2.4] Create new field “author” type people in list “CSOM Test” then migrate all list items to set user admin to field “CSOM Test Author”

                    //await CreateSiteFieldPeopleInList(ctx, LIST_NAME, "author", "author", groupName);
                    //await MigrateAllListItemsToSetUserAdmin(ctx, LIST_NAME);

                    //[3.1] Create Taxonomy Field which allow multi values, with name “cities” map to your termset.
                    //await CreateSiteFieldTypeTaxonomyMuti(ctx, "cities", "cities", groupName);

                    //[3.2] Add field “cities” to content type “CSOM Test content type” make sure don’t need update list but added field
                    //should be available in your list “CSOM test”
                    //await AddFieldToContentType(ctx, "cities", ContentTypeId);

                    //[3.3] Add 3 list item to list “CSOM test” and set multi value to field “cities” 
                    //await CreateListItem(ctx, LIST_NAME, DateTime.Today.ToString(), null, new List<string> { HO_CHI_MINH });
                    //await CreateListItem(ctx, LIST_NAME, DateTime.Today.ToString(), STOCKHOLM, new List<string> { STOCKHOLM });
                    //await CreateListItem(ctx, LIST_NAME, DateTime.Today.ToString(), STOCKHOLM, new List<string> { HO_CHI_MINH, STOCKHOLM });
                    //await CreateListItem(ctx, LIST_NAME, DateTime.Today.ToString(), HO_CHI_MINH, new List<string> { STOCKHOLM, HO_CHI_MINH });

                    //[3.4] Create new List type Document lib name “Document Test” add content type “CSOM Test content type” to this list.
                    string DOCUMENT_LIST_NAME = "Document Test";
                    //await CreateDocumentLib(ctx, DOCUMENT_LIST_NAME, "About Document Test");
                    //await AddContentTypeToList(ctx, ContentTypeName, DOCUMENT_LIST_NAME);

                    //[3.5]Create Folder “Folder 1” in root of list “Document Test”
                    //Create “Folder 2” inside “Folder 1”,
                    //Create 3 list items in “Folder 2” with value “Folder test” in field “about”.
                    //Create 2 flies in “Folder 2” with value “Stockholm” in field “cities”.

                    //await CreateFolderInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1");
                    //await CreateFolderInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2");

                    //await CreateFolderInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "Folder test");
                    //await CreateFolderInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "Folder test 1");
                    //await CreateFolderInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "Folder test 2");

                    //await CreateFileInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "File test", new List<string> { STOCKHOLM });
                    //await CreateFileInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "File test 1", new List<string> { STOCKHOLM });
                    //await CreateFileInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "File test 3", new List<string> { HO_CHI_MINH });
                    //await CreateFileInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "File test 4", new List<string> { HO_CHI_MINH, STOCKHOLM });
                    //await CreateFileInDocumnetLib(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "File test 5", new List<string> { STOCKHOLM });
                    //await CreateFileInDocumnetLibWithCT(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "File test 6", new List<string> { STOCKHOLM });
                    //await CreateFileInDocumnetLibWithCT(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "File test 7", new List<string> { STOCKHOLM });

                    //[3.6] Write CAML get all list item just in “Folder 2” and have value “Stockholm” in “cities” field
                    //await CAMLQueryWithWhere(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2");

                    //[3.7] Create List Item in “Document Test” by upload a file Document.docx 
                    //await CreateFileInDocumnetLibByUpload(ctx, DOCUMENT_LIST_NAME, "Folder 1", "Folder 2", "Document.docx");

                    //[4.1] Create View “Folders” in List “Document Test” which only show folder structure, and set this view as default
                    //await TestGetAllFolder(ctx, DOCUMENT_LIST_NAME);
                    string VIEW_NAME_ALL_FOLDERS = "All Folders";
                    //await CreateListViewFolderOnly(ctx, DOCUMENT_LIST_NAME, VIEW_NAME_ALL_FOLDERS);
                    //await SetCurrentViewAsDefault(ctx, DOCUMENT_LIST_NAME, VIEW_NAME_ALL_FOLDERS);

                    //[4.2] Write code to load User from user email or name

                    //await LoadUserFromEmailOrName(ctx, "gdfgesgfsdfs");
                    //await LoadUserFromEmailOrName(ctx, "Hân Phan Gia");
                    //await LoadUserFromEmailOrName(ctx, "GiaHan2206@y48hl.onmicrosoft.com");

                    //[4.4] tìm hiểu về TaxonomyHiddenList
                    /*
                     * Mô tả: danh sách ẩn trên web gốc.
                     *          DS chứa all thuật ngữ và từ khóa được sử dụng site
                     *          DS được dùng để lưu trữ các thuật ngữ phân loại
                     *          SharePoint dùng DS này để không cần truy vấn mà vẫn lấy đc dữ liệu meta
                     *          Khi chúng ta sử dụng term, một mục sẽ được tạo trong danh sách này. Id của mục này là WSSID. 
                     *              WSSID này thuộc loại Counter và có thể được sử dụng cho mục đích truy vấn
                     *          Đối với các term vẫn chưa được sử dụng trong trang web, không có WSSID cho thuật ngữ đó. 
                     * Công dụng:
                     * URL: https://y48hl.sharepoint.com/sites/ITcompany/Lists/TaxonomyHiddenList/AllItems.aspx
                     */
                    //[4.5] tìm hiểu về function EnsureUser và cách hoạt động
                    /*
                     * Công dụng: Kiểm tra tên đăng nhập/email có hợp lệ hay không.
                     *              Nếu tồn tại nó sẽ lấy thông tin người dùng add vô User Information List
                     *              Ta có thể dùng phương thức này để lấy User từ những thông tin người dùng cung cấp
                     * Input: string Name or Email
                     * Output: User (nếu tồn tại)
                     */

                    //Console.WriteLine(GetContentTypeByName(ctx, ContentTypeName).Id);

                    // PERMISSION EXERCISE
                    string GROUP_NAME = "tesst group";
                    string GROUP_DES = "for testing";
                    //await CreateGroup(ctx, GROUP_NAME, GROUP_DES);
                    //await AddUser(ctx, GROUP_NAME, "");

                    //--------Exercise 3 – Permission Inheritance
                    //[3.1]In the Finance and Accounting subsite, go to the List settings of the Accounts custom list and 
                    //stop inheriting permissions.
                    //Add another user to the permission list with Design permissions. 
                    //Note: You may have to refresh the page after you grant permissions

                    string SUBSITE_NAME = "Finance and Accounting";
                    string SUBSITE_URL = "FinanceAndAccounting";
                    string SUBSITE_LIST_NAME = "Accounts";
                    //await CreateSubSite(ctx, SUBSITE_URL, SUBSITE_NAME, "This is subsite");
                    //await CreateList(ctx, SUBSITE_LIST_NAME, "List in subsite", SUBSITE_NAME);
                    //await StopInheritingPermission(ctx, SUBSITE_URL, SUBSITE_LIST_NAME, true);
                    //await AddUserToPermissionList(ctx, SUBSITE_URL, SUBSITE_LIST_NAME, "test1@y48hl.onmicrosoft.com", "Design");

                    //[3.2] re-establish inheritance in list "Accounts".
                    //await ReEstablishInheritingPermission(ctx, SUBSITE_URL, SUBSITE_LIST_NAME);

                    //--------Exercise 4 – Creating	Permission	Levels	and	Groups
                    //[4.1] Get all permission group in Subsite "Finance and Accounting"
                    //await GetAllGroupPermissionLevel(ctx, SUBSITE_URL);

                    //[4.2] Create PerLev "Test Level" in Rootsite with Manage&CreateAlert
                    await CreatePermissionLevelWithManageAndCreateAlertInRoot(ctx);
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

            newList.ContentTypesEnabled = true;
            newList.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateList(ClientContext ctx, string listName, string description, string subsiteName)
        {
            Console.WriteLine("Using CSOM create a list name: " + listName);

            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = listName;
            creationInfo.Description = description;
            creationInfo.TemplateType = (int)ListTemplateType.GenericList; //Custom list

            var subWebs = ctx.Web.Webs;
            ctx.Load(subWebs);
            ctx.ExecuteQuery();

            foreach (var subWeb in subWebs)
            {
                ctx.Load(subWeb);
                await ctx.ExecuteQueryAsync();
                if (subWeb.Title == subsiteName)
                {
                    List newList = subWeb.Lists.Add(creationInfo);
                    ctx.Load(newList);
                    // Execute the query to the server.
                    await ctx.ExecuteQueryAsync();

                    newList.ContentTypesEnabled = true;
                    newList.Update();
                    await ctx.ExecuteQueryAsync();
                }
                
            }
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
        private static async Task CreateSiteFieldTypeText(ClientContext ctx, string displayName, string name, string groupName)
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
            Field field = rootWeb.Fields.AddFieldAsXml($"<Field DisplayName='{displayName}' Name='{name}' Group='{groupName}' Type='TaxonomyFieldType'/>",
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

            if (about != null)
            {
                oListItem["about"] = about;
            }

            string fieldValue;
            if (city == "Ho Chi Minh")
            {
                fieldValue = "Ho Chi Minh|90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
            }
            else if (city == "Stockholm")
            {
                fieldValue = "Stockholm|f50c5a60-1411-447d-81ca-4242f11d5380";
            }
            else
            {
                fieldValue = null;
            }

            if(fieldValue != null)
            {
                UpdateTaxonomyField(ctx, oList, oListItem, "city", fieldValue);
            }
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
            ctx.Load(listItems);
            await ctx.ExecuteQueryAsync();

            foreach (ListItem oListItem in listItems)
            {
                TaxonomyFieldValue taxFieldValue = oListItem["city"] as TaxonomyFieldValue;
                Console.WriteLine("about: {0}  - city: {1}", oListItem["about"], taxFieldValue.Label);
            }
        }
        private static async Task CreateListViewWithOrderNewestAndWhereCityInHoChiMinh(ClientContext ctx, string listName, string title)
        {
            List targetList = ctx.Web.Lists.GetByTitle(listName);

            ViewCollection viewCollection = targetList.Views;
            ctx.Load(viewCollection);

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = title;

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
            listView.Title = title;

            // You can optionally specify Aggregation: Field references for totals columns or calculated columns
            //listView.Aggregations = "<FieldRef Name='Title' Type='COUNT'/>";

            listView.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CAMLQueryUpdateMutiListItems(ClientContext ctx, string listName, string fieldName, string fielsType, string currentValue, string newValue)
        {
            List list = ctx.Web.Lists.GetByTitle(listName);

            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View>"
               + "<Query>"
               + $"<Where><Eq><FieldRef Name='{fieldName}' /><Value Type='{fielsType}'>{currentValue}</Value></Eq></Where>"
               + "</Query>"
               + "</View>";
            // execute the query
            ListItemCollection listItems = list.GetItems(query);
            ctx.Load(listItems);
            await ctx.ExecuteQueryAsync();

            foreach (ListItem oListItem in listItems)
            {
                oListItem[fieldName] = newValue;
                oListItem.Update();
                await ctx.ExecuteQueryAsync();
            }

            foreach (ListItem oListItem in listItems)
            {
                TaxonomyFieldValue taxFieldValue = oListItem["city"] as TaxonomyFieldValue;
                Console.WriteLine("about: {0}  - city: {1}", oListItem["about"], taxFieldValue.Label);
            }
        }

        private static async Task CreateSiteFieldPeople(ClientContext ctx, string displayName, string name, string groupName)
        {
            Web rootWeb = ctx.Site.RootWeb;
            // Mind the AddFieldOptions.AddFieldInternalNameHint flag
            rootWeb.Fields.AddFieldAsXml($"<Field DisplayName='{displayName}' Name='{name}' StaticName='{name}' Group='{groupName}' Type='User'/>",
                false,
                AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateSiteFieldPeopleInList(ClientContext ctx, string listName, string displayName, string name, string groupName)
        {
            List list = ctx.Web.Lists.GetByTitle(listName);
            // Mind the AddFieldOptions.AddFieldInternalNameHint flag
            list.Fields.AddFieldAsXml($"<Field DisplayName='{displayName}' Name='{name}' StaticName='{name}' Group='{groupName}' Type='User'/>",
                false,
                AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryAsync();
        }
        private static async Task MigrateAllListItemsToSetUserAdmin(ClientContext ctx, string listName)
        {
            List list = ctx.Web.Lists.GetByTitle(listName);
            User user = list.Author;

            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View>"
               + "<Query>"
               + $"<Where><Neq><FieldRef Name='city' /><Value Type='city'></Value></Eq></Where>"
               + "</Query>"
               + "</View>";
            // execute the query
            ListItemCollection listItems = list.GetItems(query);
            ctx.Load(listItems);
            await ctx.ExecuteQueryAsync();

            foreach (ListItem oListItem in listItems)
            {
                oListItem["author0"] = user;
                oListItem.Update();
                await ctx.ExecuteQueryAsync();
            }

            foreach (ListItem oListItem in listItems)
            {
                TaxonomyFieldValue taxFieldValue = oListItem["city"] as TaxonomyFieldValue;
                FieldUserValue userValue = oListItem["author0"] as FieldUserValue;
                User author = ctx.Web.EnsureUser(userValue.Email);
                ctx.Load(author);
                await ctx.ExecuteQueryAsync();
                Console.WriteLine("about: {0}  - city: {1} - author: {2}", oListItem["about"], taxFieldValue.Label, author.Title);
            }
        }

        //Exercise 3:
        private static async Task CreateSiteFieldTypeTaxonomyMuti(ClientContext ctx, string displayName, string name, string groupName)
        {
            Web rootWeb = ctx.Site.RootWeb;
            Field field = rootWeb.Fields.AddFieldAsXml($"<Field DisplayName='{displayName}' Name='{name}' Group='{groupName}' Type='TaxonomyFieldTypeMulti'/>",
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
            taxonomyField.AllowMultipleValues = true;
            taxonomyField.Update();

            await ctx.ExecuteQueryAsync();
        }
        public static void UpdateTaxonomyFieldMulti(ClientContext ctx, List list, ListItem listItem, string fieldName, string fieldValue)
        {
            Field field = list.Fields.GetByInternalNameOrTitle(fieldName);
            TaxonomyField txField = ctx.CastTo<TaxonomyField>(field);
            TaxonomyFieldValueCollection termValue = new TaxonomyFieldValueCollection(
                ctx,
                fieldValue,
                txField);
            txField.SetFieldValueByValueCollection(listItem, termValue);
            listItem.Update();
            ctx.Load(listItem);
            ctx.ExecuteQuery();
        }
        private static async Task CreateListItem(ClientContext ctx, string listName, string about, string city, List<string> cities)
        {
            List oList = ctx.Web.Lists.GetByTitle(listName);

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = oList.AddItem(itemCreateInfo);

            if (about != null)
            {
                oListItem["about"] = about;
            }
            string fieldValue = "-1;";
            int count = 0;
            foreach (string city1 in cities)
            {
                if (count != 0)
                {
                    fieldValue += ";#-1;";
                }
                if (city1 == "Ho Chi Minh")
                {
                    fieldValue += "#Ho Chi Minh|90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
                }
                else if (city1 == "Stockholm")
                {
                    fieldValue += "#Stockholm|f50c5a60-1411-447d-81ca-4242f11d5380";
                }
                count++;
            }
            UpdateTaxonomyFieldMulti(ctx, oList, oListItem, "cities", fieldValue);

            string fieldValue1;
            if (city == "Ho Chi Minh")
            {
                fieldValue1 = "Ho Chi Minh|90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
            }
            else if (city == "Stockholm")
            {
                fieldValue1 = "Stockholm|f50c5a60-1411-447d-81ca-4242f11d5380";
            }
            else
            {
                fieldValue1 = null;
            }

            if (fieldValue1 != null)
            {
                UpdateTaxonomyField(ctx, oList, oListItem, "city", fieldValue1);
            }
            oListItem.Update();
            await ctx.ExecuteQueryAsync();

        }
        private static async Task CreateDocumentLib(ClientContext ctx, string documentLibName, string description)
        {
            Console.WriteLine("Using CSOM create a documnet libary name: " + documentLibName);

            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = documentLibName;
            creationInfo.Description = description;
            creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary; //Custom list

            List newList = ctx.Web.Lists.Add(creationInfo);
            ctx.Load(newList);
            // Execute the query to the server.
            await ctx.ExecuteQueryAsync();

            newList.ContentTypesEnabled = true;
            newList.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateFolderInDocumnetLib(ClientContext ctx, string documentLibName, string folderName)
        {
            List list = ctx.Web.Lists.GetByTitle(documentLibName);
            list.EnableFolderCreation = true;
            list.Update();
            await ctx.ExecuteQueryAsync();

            //To create the folder
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

            itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
            itemCreateInfo.LeafName = folderName;

            ListItem newItem = list.AddItem(itemCreateInfo);
            newItem["Title"] = folderName;
            newItem.Update();
            ctx.ExecuteQuery();
        }
        private static async Task CreateFolderInDocumnetLib(ClientContext ctx, string documentLibName, string rootFolderName, string folderName)
        {
            List list = ctx.Web.Lists.GetByTitle(documentLibName);
            list.EnableFolderCreation = true;
            list.Update();
            await ctx.ExecuteQueryAsync();

            //To create the folder
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

            itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
            itemCreateInfo.LeafName = folderName;

            FolderCollection folders = list.RootFolder.Folders;
            ctx.Load(folders);
            ctx.ExecuteQuery();
            foreach (Microsoft.SharePoint.Client.Folder folder in folders)
            {
                Console.WriteLine(folder.Name);

                if (folder.Name == rootFolderName)
                {
                    folder.Folders.Add(folderName);
                }
            }
            await ctx.ExecuteQueryAsync();

        }
        private static async Task CreateFolderInDocumnetLib(ClientContext ctx, string documentLibName, string rootFolderName, string subFolderName, string folderName)
        {
            List list = ctx.Web.Lists.GetByTitle(documentLibName);
            list.EnableFolderCreation = true;
            list.Update();
            await ctx.ExecuteQueryAsync();

            FolderCollection folders = list.RootFolder.Folders;
            ctx.Load(folders);
            ctx.ExecuteQuery();
            foreach (Microsoft.SharePoint.Client.Folder folder in folders)
            {
                if (folder.Name == rootFolderName)
                {
                    FolderCollection folders1 = folder.Folders;
                    ctx.Load(folders1);
                    await ctx.ExecuteQueryAsync();
                    foreach (Microsoft.SharePoint.Client.Folder folder1 in folders1)
                    {
                        if (folder1.Name == subFolderName)
                        {
                            Console.WriteLine(folder.Name + "/" + folder1.Name);
                            folder1.Folders.Add(folderName);
                            await ctx.ExecuteQueryAsync();
                            //ListItem listItem = folder1.ListItemAllFields;
                            //listItem["about"] = "test";
                            //listItem.Update();
                            //await ctx.ExecuteQueryAsync();
                            FolderCollection folders2 = folder1.Folders;
                            ctx.Load(folders2);
                            await ctx.ExecuteQueryAsync();
                            foreach (Microsoft.SharePoint.Client.Folder folder2 in folders2)
                            {
                                if (folder2.Name == folderName)
                                {
                                    Console.WriteLine(folder.Name + "/" + folder1.Name + "/" + folder2.Name);
                                    ListItem listItem = folder2.ListItemAllFields;
                                    listItem["about"] = "test";
                                    listItem.Update();
                                    await ctx.ExecuteQueryAsync();
                                }
                            }
                        }
                    }
                }
            }
        }
        private static async Task Test(ClientContext ctx)
        {
            Microsoft.SharePoint.Client.Folder targetFolder = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + "Document Test/Folder 1/Folder 2");
            ctx.Load(targetFolder);
            ctx.ExecuteQuery();

            Microsoft.SharePoint.Client.Folder newFolder = targetFolder.Folders.Add("Folder Test 22");
            ListItem item = newFolder.ListItemAllFields;
            item["about"] = "Folder Test";
            item.Update();
            await ctx.ExecuteQueryAsync();

        }
        private static async Task CreateFileInDocumnetLib(ClientContext ctx, string documentLibName, string rootFolderName, string subFolderName, string fileName, List<string> cities)
        {
            Microsoft.SharePoint.Client.Folder targetFolder = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + documentLibName + "/" + rootFolderName + "/" + subFolderName);
            ctx.Load(targetFolder);
            ctx.ExecuteQuery();

            FileCreationInformation createFile = new FileCreationInformation();
            createFile.Url = $"{fileName}.txt";
            //use byte array to set content of the file
            string somestring = "hello there";
            byte[] toBytes = Encoding.ASCII.GetBytes(somestring);

            createFile.Content = toBytes;

            Microsoft.SharePoint.Client.File newFile = targetFolder.Files.Add(createFile);
            ctx.Load(newFile);
            await ctx.ExecuteQueryAsync();

            //UPDATE FIELD CITIES
            ListItem item = newFile.ListItemAllFields;

            string fieldValue = "-1;";
            int count = 0;
            foreach (string city in cities)
            {
                if (count != 0)
                {
                    fieldValue += ";#-1;";
                }
                if (city == "Ho Chi Minh")
                {
                    fieldValue += "#Ho Chi Minh|90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
                }
                else if (city == "Stockholm")
                {
                    fieldValue += "#Stockholm|f50c5a60-1411-447d-81ca-4242f11d5380";
                }
                count++;
            }

            Field field = ctx.Web.Fields.GetByInternalNameOrTitle("cities");
            TaxonomyField txField = ctx.CastTo<TaxonomyField>(field);
            TaxonomyFieldValueCollection termValue = new TaxonomyFieldValueCollection(
                ctx,
                fieldValue,
                txField);
            txField.SetFieldValueByValueCollection(item, termValue);
            item.Update();
            ctx.Load(item);
            await ctx.ExecuteQueryAsync();

            item.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateFileInDocumnetLibWithCT(ClientContext ctx, string documentLibName, string rootFolderName, string subFolderName, string fileName, List<string> cities)
        {
            Microsoft.SharePoint.Client.Folder targetFolder = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + documentLibName + "/" + rootFolderName + "/" + subFolderName);
            ctx.Load(targetFolder);
            ctx.ExecuteQuery();

            FileCreationInformation createFile = new FileCreationInformation();
            createFile.Url = $"{fileName}.txt";
            //use byte array to set content of the file
            string somestring = "hello there";
            byte[] toBytes = Encoding.ASCII.GetBytes(somestring);

            createFile.Content = toBytes;

            Microsoft.SharePoint.Client.File newFile = targetFolder.Files.Add(createFile);
            ctx.Load(newFile);
            await ctx.ExecuteQueryAsync();

            //UPDATE FIELD CITIES
            ListItem item = newFile.ListItemAllFields;

            string fieldValue = "-1;";
            int count = 0;
            foreach (string city in cities)
            {
                if (count != 0)
                {
                    fieldValue += ";#-1;";
                }
                if (city == "Ho Chi Minh")
                {
                    fieldValue += "#Ho Chi Minh|90dd8af9-e9f0-4f6e-ac57-68200c8ea34c";
                }
                else if (city == "Stockholm")
                {
                    fieldValue += "#Stockholm|f50c5a60-1411-447d-81ca-4242f11d5380";
                }
                count++;
            }

            Field field = ctx.Web.Fields.GetByInternalNameOrTitle("cities");
            TaxonomyField txField = ctx.CastTo<TaxonomyField>(field);
            TaxonomyFieldValueCollection termValue = new TaxonomyFieldValueCollection(
                ctx,
                fieldValue,
                txField);
            txField.SetFieldValueByValueCollection(item, termValue);

            item["ContentTypeId"] = "0x0101009189AB5D3D2647B580F011DA2F356FB2";
            item.Update();
            ctx.Load(item);
            await ctx.ExecuteQueryAsync();

            item.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CAMLQueryWithWhere(ClientContext ctx, string listName, string folderLv1, string folderLv2)
        {
            Microsoft.SharePoint.Client.Folder targetFolder = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + listName + "/" + folderLv1 + "/" + folderLv2);
            ctx.Load(targetFolder);
            ctx.ExecuteQuery();

            List list = ctx.Web.Lists.GetByTitle(listName);
            ctx.Load(list);
            ctx.ExecuteQuery();

            var results = new Dictionary<string, IEnumerable<Microsoft.SharePoint.Client.File>>();
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View Scope='RecursiveAll'>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='cities' />
                                            <Value Type='Text'>Stockholm</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                            </View>";
            // execute the query

            query.FolderServerRelativeUrl = targetFolder.ServerRelativeUrl;
            ListItemCollection listItems = list.GetItems(query);

            ctx.Load(listItems, icol => icol.Include(i => i.File));
            var a = listItems.Select(i => i.File);
            await ctx.ExecuteQueryAsync();

           // foreach (var result in results)
            {
                foreach (var file in a)
                {

                    Console.WriteLine("File: {0}", file.Name);
                }
            }
        }
        private static async Task CreateFileInDocumnetLibByUpload(ClientContext ctx, string documentLibName, string rootFolderName, string subFolderName, string folderUploadUrl)
        {
            Microsoft.SharePoint.Client.Folder targetFolder = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + documentLibName + "/" + rootFolderName + "/" + subFolderName);
            ctx.Load(targetFolder);
            ctx.ExecuteQuery();

            FileCreationInformation createFile = new FileCreationInformation();
            createFile.Content = System.IO.File.ReadAllBytes(folderUploadUrl);
            createFile.Overwrite = true;
            createFile.Url = Path.GetFileName(folderUploadUrl)
;
            Microsoft.SharePoint.Client.File newFile = targetFolder.Files.Add(createFile);
            ctx.Load(newFile);
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateListViewFolderOnly(ClientContext ctx, string listName, string viewName)
        {
            List targetList = ctx.Web.Lists.GetByTitle(listName);

            ViewCollection viewCollection = targetList.Views;
            ctx.Load(viewCollection);

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = viewName;

            // Specify type of the view. Below are the options

            // 1. none - The type of the list view is not specified
            // 2. html - Sspecifies an HTML list view type
            // 3. grid - Specifies a datasheet list view type
            // 4. calendar- Specifies a calendar list view type
            // 5. recurrence - Specifies a list view type that displays recurring events
            // 6. chart - Specifies a chart list view type
            // 7. gantt - Specifies a Gantt chart list view type

            viewCreationInformation.ViewTypeKind = ViewType.Html;

            // You can optionally specify row limit for the view
            //viewCreationInformation.RowLimit = 10;

            // You can optionally specify a query as mentioned below.
            // Create one CAML query to filter list view and mention that query below
            viewCreationInformation.Query =
                    "    <Where>"
                  + "      <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq>"
                  + "    </Where>";
            //viewCreationInformation.Query = "<OrderBy><FieldRef Name='Created' Ascending='FALSE'/></OrderBy>";

            // Add all the fields over here with comma separated value as mentioned below
            // You can mention display name or internal name of the column
            string CommaSeparateColumnNames = "Name,about,city";
            viewCreationInformation.ViewFields = CommaSeparateColumnNames.Split(',');

            View listView = viewCollection.Add(viewCreationInformation);
            ctx.ExecuteQuery();

            // Code to update the display name for the view.
            listView.Title = viewName;

            // You can optionally specify Aggregation: Field references for totals columns or calculated columns
            //listView.Aggregations = "<FieldRef Name='Title' Type='COUNT'/>";

            listView.Update();
            await ctx.ExecuteQueryAsync();
        }
        private static async Task TestGetAllFolder(ClientContext clientContext, string title)
        {
            List spList = clientContext.Web.Lists.GetByTitle(title);
            clientContext.Load(spList);
            clientContext.ExecuteQuery();

            if (spList != null && spList.ItemCount > 0)
            {
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml =
                    "<View Scope='RecursiveAll'>"
                  + "  <Query>"
                  + "    <Where>"
                  + "      <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq>"
                  + "    </Where>"
                  + "  </Query>"
                  + "  <ViewFields><FieldRef Name='Title' /></ViewFields>"
                  + "</View>";

                ListItemCollection listItems = spList.GetItems(camlQuery);

                clientContext.Load(listItems);
                await clientContext.ExecuteQueryAsync();

                foreach (var item in listItems)
                {
                    //Console.WriteLine($"Title: {item.FieldValues["Title"]} - FileRef: { item.FieldValues["FileRef"]}-FileLeafRef: { item.FieldValues["FileLeafRef"]}");
                    Console.WriteLine($"Title: {item.FieldValues["FileRef"]}");
                }
            }
        }
        private static async Task SetCurrentViewAsDefault(ClientContext ctx, string listName,string viewName)
        {
            List list = ctx.Web.Lists.GetByTitle(listName);
            View view = list.Views.GetByTitle(viewName);

            view.DefaultView = true;
            view.Update();

            await ctx.ExecuteQueryAsync();  
        }
        private static async Task LoadUserFromEmailOrName(ClientContext ctx, string nameOrMail)
        {
            try
            {
                User user = ctx.Web.EnsureUser(nameOrMail);
                ctx.Load(user);
                await ctx.ExecuteQueryAsync();

                Console.WriteLine(user.Email);
                Console.WriteLine(user.Title);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Not a valid person");
            }
            Console.ReadLine();
        }
        private static ContentType GetContentTypeByName(ClientContext ctx, string contentTypeName)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();

            return Enumerable.FirstOrDefault(contentTypes, ct => ct.Name == contentTypeName);
        }
        
        //exercise permission
        private static async Task CreateSubSite(ClientContext ctx, string subsiteURL, string subsiteTitle, string subsiteDes)
        {
            WebCreationInformation oWebCreationInformation = new WebCreationInformation();
            // This is relative URL of the url provided in context
            oWebCreationInformation.Url = subsiteURL;
            oWebCreationInformation.Title = subsiteTitle;
            oWebCreationInformation.Description = subsiteDes;

            // This will inherit permission from parent site
            oWebCreationInformation.UseSamePermissionsAsParentSite = true;

            // "STS#0" is the code for 'Team Site' template
            oWebCreationInformation.WebTemplate = "STS#0";
            oWebCreationInformation.Language = 1033;

            Web oWeb = ctx.Site.RootWeb.Webs.Add(oWebCreationInformation);
            await ctx.ExecuteQueryAsync();
        }
        private static async Task StopInheritingPermission(ClientContext ctx, string subsiteurl,string listName, bool copyRoleAssigements)
        {
            Web subsite = ctx.Site.OpenWeb(ctx.Web.ServerRelativeUrl +"/" + subsiteurl);
            Console.WriteLine(ctx.Web.ServerRelativeUrl + "/" + subsiteurl);
            List list = subsite.Lists.GetByTitle(listName);
            ctx.Load(list);
            ctx.ExecuteQuery();

            list.BreakRoleInheritance(copyRoleAssigements, false);

            await ctx.ExecuteQueryAsync();
        }
        private static async Task ReEstablishInheritingPermission(ClientContext ctx, string subsiteurl, string listName)
        {
            Web subsite = ctx.Site.OpenWeb(ctx.Web.ServerRelativeUrl + "/" + subsiteurl);
            List list = subsite.Lists.GetByTitle(listName);
            ctx.Load(list);
            ctx.ExecuteQuery();

            list.ResetRoleInheritance();

            await ctx.ExecuteQueryAsync();
        }
        private static async Task AddUserToPermissionList(ClientContext ctx, string subsiteurl, string listName, string loginName, string perLv)
        {
            Web subsite = ctx.Site.OpenWeb(ctx.Web.ServerRelativeUrl + "/" + subsiteurl);
            Console.WriteLine(ctx.Web.ServerRelativeUrl + "/" + subsiteurl);
            List list = subsite.Lists.GetByTitle(listName);
            ctx.Load(list);
            ctx.ExecuteQuery();

            try
            {
                User u = ctx.Web.EnsureUser(loginName);
                ctx.Load(u);
                await ctx.ExecuteQueryAsync();

                Console.WriteLine(u.LoginName);

                Principal user = ctx.Web.SiteUsers.GetByLoginName(u.LoginName);

                RoleDefinition writeDefinition = subsite.RoleDefinitions.GetByName(perLv);
                RoleDefinitionBindingCollection roleDefCollection = new RoleDefinitionBindingCollection(ctx);
                roleDefCollection.Add(writeDefinition);
                RoleAssignment newRoleAssignment = list.RoleAssignments.Add(user, roleDefCollection);

                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        private static async Task GetAllGroupPermissionLevel(ClientContext ctx, string subsiteurl)
        {
            Web subsite = ctx.Site.OpenWeb(ctx.Web.ServerRelativeUrl + "/" + subsiteurl);

            RoleDefinitionCollection roleDefinitions = subsite.RoleDefinitions;
            ctx.Load(roleDefinitions);
            ctx.ExecuteQuery();

            foreach (RoleDefinition roleDefinition in roleDefinitions)
            {
                Console.WriteLine(roleDefinition.Name);
            }

            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateGroup(ClientContext ctx, string groupName, string groupDescription)
        {
            GroupCollection oGroupCollection = ctx.Web.SiteGroups;
            // GroupCreationInformation object
            GroupCreationInformation oGroupCreationInformation = new GroupCreationInformation();

            oGroupCreationInformation.Title = groupName;
            oGroupCreationInformation.Description = groupDescription;

            Group oGroup = oGroupCollection.Add(oGroupCreationInformation);
            await ctx.ExecuteQueryAsync();
        }
        private static async Task AddUser(ClientContext ctx, string groupName, string loginNameOrEmail)
        {
            GroupCollection groups = ctx.Web.SiteGroups;
            Group group = groups.GetByName(groupName); 
            ctx.Load(group);
            ctx.ExecuteQuery();

            Console.WriteLine(group.Title);

            User aoUser = ctx.Web.EnsureUser(loginNameOrEmail);
            User oUser = group.Users.AddUser(aoUser);
            ctx.ExecuteQuery();
        }
        private static async Task CreatePermissionLevelWithManageAndCreateAlertInRoot(ClientContext ctx)
        {
            Web web = ctx.Web;
            ctx.Load(web);
            ctx.Load(web.AllProperties);
            ctx.Load(web.RoleDefinitions);
            ctx.ExecuteQuery();
            var roleDefinitions = web.RoleDefinitions;

            BasePermissions basePermissions = new BasePermissions();
            basePermissions.Set(PermissionKind.ManageLists);
            basePermissions.Set(PermissionKind.CreateAlerts);

            RoleDefinitionCreationInformation roleDefinitionCreationInfo = new RoleDefinitionCreationInformation();
            roleDefinitionCreationInfo.BasePermissions = basePermissions;
            roleDefinitionCreationInfo.Name = "Test Level";
            roleDefinitionCreationInfo.Description = "create alert and manage list";
            RoleDefinition roleDefinition = ctx.Web.RoleDefinitions.Add(roleDefinitionCreationInfo);
            await ctx.ExecuteQueryAsync();
        }
    }
}
