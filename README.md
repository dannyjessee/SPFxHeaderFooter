## SharePoint Framework custom header and footer application customizer extension

This is the code for the SharePoint Framework application customizer extension that renders a custom header and footer on all modern pages within a SharePoint Online site. You can read more details about it in my blog post [here](https://dannyjessee.com/blog/index.php/2017/08/custom-modern-page-header-and-footer-using-sharepoint-framework-part-2/).

![Modern page header and footer](https://i1.wp.com/dannyjessee.com/blog/wp-content/uploads/2017/08/modernhf2.png?w=953&ssl=1)

Note that this solution complements the [SharePoint-hosted add-in](https://github.com/dannyjessee/SiteHeaderFooter) that includes the configuration interface and user custom actions to render the header and footer on all classic pages within a SharePoint Online site. You can read more details about that [here](https://dannyjessee.com/blog/index.php/2015/08/custom-site-header-and-footer-using-a-sharepoint-hosted-add-in/). <b>To use this application customizer extension, you will need to first install and configure the add-in.</b>

![Add-in part and classic page header and footer](https://i1.wp.com/dannyjessee.com/blog/wp-content/uploads/2017/08/classicaddinsuitebar.png?w=953&ssl=1)

### Building the code (for localhost debugging)

```bash
git clone the repo
npm i
gulp serve --nobrowser
```

When debugging locally, you will need to include the following querystring parameter on any modern pages where you wish to test this functionality:

```bash
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"bbe5f3fa-7326-455d-8573-9f0b2b015ff9":{"location":"ClientSideExtension.ApplicationCustomizer"}}
```

### Building the code (for production deployment)

```bash
git clone the repo
npm i
gulp bundle --ship
gulp package-solution --ship
```

### Tenant-scoped deployment

This extension is configured to optionally allow [tenant-scoped deployment](https://dev.office.com/sharepoint/docs/spfx/tenant-scoped-deployment). When deploying to production, you must first upload the **.sppkg** file to your tenant's app catalog. You will then be given the option to make this solution available to all sites within your organization:

![Tenant-scoped deployment](https://i1.wp.com/dannyjessee.com/blog/wp-content/uploads/2017/09/tenantscopeddeployment.png?w=784&ssl=1)

### Adding the user custom action for the extension in a tenant-scoped deployment

If you check the box labeled **Make this solution available to all sites in the organization** before pressing **Deploy**, you will need to manually add the user custom action associated with the extension on any site where you would like the custom header and footer to be rendered on modern pages. If you deploy the extension at tenant scope, it is immediately available to all sites and you do not need to explicitly add the app from the Site Contents screen. However, because tenant-scoped extensions cannot leverage the feature framework, you will need to associate the user custom action with the **ClientSideComponentId** of the extension manually. This can be accomplished a number of different ways. Some example code using the .NET Managed Client Object Model in a console application is shown below:

```cs
using (ClientContext ctx = new ClientContext("https://[YOUR TENANT].sharepoint.com"))
{
    SecureString password = new SecureString();
    foreach (char c in "[YOUR PASSWORD]".ToCharArray()) password.AppendChar(c);
    ctx.Credentials = new SharePointOnlineCredentials("[USER]@[YOUR TENANT].onmicrosoft.com", password);

    Web web = ctx.Web;
    UserCustomActionCollection ucaCollection = web.UserCustomActions;
    UserCustomAction uca = ucaCollection.Add();
    uca.Title = "SPFxHeaderFooterApplicationCustomizer";
    // This is the user custom action location for application customizer extensions
    uca.Location = "ClientSideExtension.ApplicationCustomizer";
    // Use the ID from HeaderFooterApplicationCustomizer.manifest.json below
    uca.ClientSideComponentId = new Guid("bbe5f3fa-7326-455d-8573-9f0b2b015ff9");
    uca.Update();

    ctx.Load(web, w => w.UserCustomActions);
    ctx.ExecuteQuery();

    Console.WriteLine("User custom action added to site successfully!");
}
```

You still need to install and configure the [SharePoint-hosted add-in](https://github.com/dannyjessee/SiteHeaderFooter) on any site where you wish to use this extension and [disable NoScript on that site](https://dannyjessee.com/blog/index.php/2017/07/sharepoint-online-modern-team-sites-are-noscript-sites-but-communication-sites-are-not/) if necessary.

### If you do not perform a tenant-scoped installation

If you decline to check the box to allow tenant-scoped installation when you upload the **.sppkg** file to the app catalog, the extension will be made available to manually add to any site via **Site Contents > Add an app**. This will automatically associate the user custom action on any site where you manually add the extension, so no code is necessary to register the user custom action as shown above in this case.
