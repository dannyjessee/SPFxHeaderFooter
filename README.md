## SharePoint Framework custom header and footer application customizer extension

This is where you include your WebPart documentation.

### Building the code (for localhost debugging)

```bash
git clone the repo
npm i
gulp serve --nobrowser
```

### Building the code (for production deployment)

```bash
git clone the repo
npm i
gulp bundle --ship
gulp package-solution --ship
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp serve --nobrowser - localhost debugging
gulp package-solution --ship - production deployment
