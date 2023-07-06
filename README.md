# Iteration Work Items Hub

## Install
Install from Visual Studio Marketplace: https://marketplace.visualstudio.com/items?itemName=JamesSkemp.iteration-work-items-hub

## Build
`npm run build` automatically increments the version and creates a new *.vsix file in the root directory.

Publish to https://marketplace.visualstudio.com/manage/publishers

## Development
Based upon the [Azure DevOps Web Sample Extension](https://github.com/microsoft/azure-devops-extension-sample).

Original README contents are below.

### Azure DevOps Web Sample Extension

This repository generates an [Azure DevOps extension](https://docs.microsoft.com/en-us/azure/devops/extend/overview?view=vsts) containing a number of different contributions of various types.

#### Dependencies

The sample repository depends on a few Azure DevOps packages:

- [azure-devops-extension-sdk](https://github.com/Microsoft/azure-devops-extension-sdk): Required module for Azure DevOps extensions which allows communication between the host page and the extension iframe.
- [azure-devops-extension-api](https://github.com/Microsoft/azure-devops-extension-api): Contains REST client libraries for the various Azure DevOps feature areas.
- [azure-devops-ui](https://developer.microsoft.com/azure-devops): UI library containing the React components used in the Azure DevOps web UI.

Some external dependencies:

- `React` - Is used to render the UI in the samples, and is a dependency of `azure-devops-ui`.
- `TypeScript` - Samples are written in TypeScript and compiled to JavaScript
- `SASS` - Extension samples are styled using SASS (which is compiled to CSS and delivered in webpack js bundles).
- `webpack` - Is used to gather dependencies into a single javascript bundle for each sample.

#### Building the sample project

Just run:

    npm run build

This produces a .vsix file which can be uploaded to the [Visual Studio Marketplace](https://marketplace.visualstudio.com/azuredevops)

#### Using the extension

The preferred way to get started is to use the `tfx extension init` command which will clone from this sample and prompt you for replacement information (like your publisher id). Just run:

    npm install -g tfx-cli
    tfx extension init

You can also clone the sample project and change the `publisher` property in `azure-devops-extension.json` to your own Marketplace publisher id. Refer to the online [documentation](https://docs.microsoft.com/en-us/azure/devops/extend/publish/overview?view=vsts) for setting up your own publisher and publishing an extension.

### Samples

Individual sample contributions are self-contained folders under `./src/Samples`. Within each sample you will find:

1. `{SampleName}.json` - describes the contribution objects being added to Azure DevOps
2. `{SampleName}.html` - page which is rendered within an iframe on the appropriate Azure DevOps page or pages. It may be visible UI (such as a Hub) or a background iframe (such as a Menu action handler). This will include a sample reference for `{SampleName}.js`, and for visible frames it will contain a single `<div>` element with an id of `root`.
3. `{SampleName}.ts(x)` - Root script that is run when the frame is loaded. A webpack entry is added for this file which will generate a single `js` file with this content and all its dependencies.
4. `{SampleName}.scss` - optional sass file containing the styles (CSS) for the UI
5. Additional ts/tsx files - For samples that are too big for one file, the code will be broken up appropriately

#### Hub

This sample adds a hub named "Sample Hub" into the `Pipelines` hub group. If you visit a project-level page, you will find Sample Hub under the `Pipelines` navigation element in the vertical navigation menu on the left of the page.

The hub uses a Pivot component to draw 4 different tabs:

1. An `Overview` tab contains some simple details about the current user and project
2. A `Navigation` tab contains a few actions that allow you to integrate with the page's URL and title
3. An `Extension Data` tab demonstrates reading and writing to the extension data service
4. A `Messages` tab shows how to display global messages

There are also actions at the top-right of the hub which demonstrate opening dialogs and panels, including custom content within them (used in the `Panel` sample).

#### QueryParamsHandler

This sample adds a service that gets loaded on any page whenever a "showMyPanel" query parameter is present
in the URL when any page is loaded. The startup service shows the custom panel from the Panel sample, using
an optional "myPanelTitle" query parameter as the panel title.

### References

The full set of documentation for developing extensions can be found at [https://docs.microsoft.com/en-us/azure/devops/extend](https://docs.microsoft.com/en-us/azure/devops/extend/?view=vsts).
