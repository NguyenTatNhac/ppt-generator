# PPT Generator App

PPT Generator is a Jira app used to export Jira issue data to a custom PPTX template.

## Building The App

### Prerequisite

* Java is installed
* Atlassian Plugin SDK is installed (
  See [Setup the Atlassian Plugin SDK](https://developer.atlassian.com/server/framework/atlassian-sdk/set-up-the-atlassian-plugin-sdk-and-build-a-project/))

### Setting up IntelliJ IDEA

The IntelliJ IDEA will use the bundled maven for the build process. In order to start working with
Jira plugin, we need to re-config, so the IDE will use the maven bundle from the Atlassian SDK.

1. Press `Ctrl+Alt+S` to open IDE settings
2. Select `Build, Execution, Deployment | Build Tools | Maven`
3. Change `Maven home path` to the maven directory from the Atlassian SDK installation folder. For example:  
   `C:\Applications\Atlassian\atlassian-plugin-sdk-8.2.7\apache-maven-3.5.4`
4. Override the `User settings file` with the `settings.xml` from the SDK bundled maven. For example:  
   `C:\Applications\Atlassian\atlassian-plugin-sdk-8.2.7\apache-maven-3.5.4\conf\settings.xml`

### Build app

To build the app for production, run this command at the project root directory:

```cmd
atlas-clean && atlas-package
```

A plugin `jar` file will be built in the `target` directory, use this file to install in the Jira
Production.

## Field Naming Note

This app will get Jira Custom Field by its name. These names can be different between test and
production instance. You can correct them in the file `com.viz.jira.app.ppt.sdo.CustomFieldName`.
After the correction, rebuild the app by running `atlas-clean && atlas-package`.
