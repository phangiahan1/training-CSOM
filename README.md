# ConsoleCSOM
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualBasic.FileIO;

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

                    //ListCreationInformation creationInfo = new ListCreationInformation();
                    //creationInfo.Title = "CSOM Test";
                    //creationInfo.Description = "using CSOM create a list";
                    //creationInfo.TemplateType = (int)ListTemplateType.GenericList; //Custom list

                    //List newList = ctx.Web.Lists.Add(creationInfo);
                    //ctx.Load(newList);
                    //// Execute the query to the server.
                    //ctx.ExecuteQuery();

                    //[1.2] Create term set "city-han" in dev tenant
                    //[1.3] Create 2 term "Ho Chi Minh" and "Stockholm" in termset "city-han"

                    //string termGroupName = "city";
                    //string termSetName = "city-han";
                    //TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                    //TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                    //if (termStore != null)
                    //{
                    //    TermGroup termGroup = termStore.CreateGroup(termGroupName, Guid.NewGuid());
                    //    TermSet myTermSet = termGroup.CreateTermSet(termSetName, Guid.NewGuid(), 1033);
                    //    myTermSet.CreateTerm("Ho Chi Minh", 1033, Guid.NewGuid());
                    //    myTermSet.CreateTerm("Stockholm", 1033, Guid.NewGuid());
                    //    ctx.ExecuteQuery();
                    //}

                    //[1.4] Create site fields "about" type text and field "city" type taxonomy

                    //Web rootWeb = ctx.Site.RootWeb;
                    //// Mind the AddFieldOptions.AddFieldInternalNameHint flag
                    //rootWeb.Fields.AddFieldAsXml("<Field DisplayName='about' " +
                    //    "Name='about' ID='{2d9c2efe-58f2-4003-85ce-0251eb174096}' " +
                    //    "Group='CSOM city projects' " +
                    //    "Type='Text' />", 
                    //    false, 
                    //    AddFieldOptions.AddFieldInternalNameHint);
                    //Field field = rootWeb.Fields.AddFieldAsXml("<Field DisplayName='city' " +
                    //    "Name='city' ID='{abf2bde8-f99b-4f76-89d0-1cb5f19695b8}' " +
                    //    "Group='CSOM city projects' Type='TaxonomyFieldType' />", 
                    //    false, 
                    //    AddFieldOptions.AddFieldInternalNameHint);

                    //ctx.ExecuteQuery();

                    //Guid termStoreId = Guid.Empty;
                    //Guid termSetId = Guid.Empty;
                    //GetTaxonomyFieldInfo(ctx, out termStoreId, out termSetId);

                    //// Retrieve as Taxonomy Field
                    //TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
                    //taxonomyField.SspId = termStoreId;
                    //taxonomyField.TermSetId = termSetId;
                    //taxonomyField.TargetTemplate = String.Empty;
                    //taxonomyField.AnchorId = Guid.Empty;
                    //taxonomyField.Update();

                    //ctx.ExecuteQuery();

                    //[1.5] Create site content type "CSOM Test content type"
                    //      => add this to "CSOM test" add fields "about" and "city" to this.

                    Web rootWeb = ctx.Site.RootWeb;

                    // create by ID
                    rootWeb.ContentTypes.Add(new ContentTypeCreationInformation
                    {
                        Name = "CSOM Test content type",
                        Id = "0x0100BDD5E43587AF469CA722FD068065DF5D",
                        Group = "CSOM city projects Content Types"
                    });
                    ctx.ExecuteQuery();

                    //[1.6] In list "CSOM test" set "CSOM Test content type" as default content type


                    //[1.7] Create 5 list items to list with some value  in field "about" and "city"


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

        private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx)
        {
            var field = ctx.Web.Fields.GetByTitle("fieldname");

            ctx.Load(field);
            await ctx.ExecuteQueryAsync();

            var taxField = ctx.CastTo<TaxonomyField>(field);

            taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "correct label here",
                TermGuid = "term id"
            });
            item.Update();
            await ctx.ExecuteQueryAsync();
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
            TermSetCollection termSets = termStore.GetTermSetsByName("SPSNL14", 1033);

            clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            clientContext.Load(termStore, ts => ts.Id);
            clientContext.ExecuteQuery();

            termStoreId = termStore.Id;
            termSetId = termSets.FirstOrDefault()!.Id;
        }
    }
}
