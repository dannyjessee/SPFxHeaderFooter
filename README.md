## SharePoint Framework custom header and footer application customizer extension

This is the code for the SharePoint Framework application customizer extension that renders a custom header and footer on all modern pages within a SharePoint Online site. You can read more details about it in my blog post [here](https://dannyjessee.com/blog/index.php/2017/06/custom-modern-page-header-and-footer-using-sharepoint-framework/).

![Modern page header and footer](https://i0.wp.com/dannyjessee.com/blog/wp-content/uploads/2017/06/contentareas3.png?w=820&ssl=1)

Note that this solution complements the SharePoint-hosted add-in that includes the configuration interface and user custom actions to render the header and footer on all classic pages within a SharePoint Online site. You can read more details about that [here](https://dannyjessee.com/blog/index.php/2015/08/custom-site-header-and-footer-using-a-sharepoint-hosted-add-in/). <b>To use this application customizer extension, you will need to first install and configure the add-in.</b>

![Add-in part and classic page header and footer](https://i1.wp.com/dannyjessee.com/blog/wp-content/uploads/2017/06/addinpartclassicpage.png?w=820&ssl=1)

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

When deploying to production, you must first upload the <b>.sppkg</b> file to your tenant's app catalog. You may then install the app from the Site Contents page within your SharePoint Online site.
