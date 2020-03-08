# excel-aws-add-in-visualstudio-code

## Prerequisites
- Node.js is to be installed.
- install generator.
  - npm install -g yo generator-office

## Develop

**Step 1:**
- $ yo office
  - ? Choose a project type: _[Office Add-in Task Pane project]_
  - ? Choose a script type: _[JavaScript]_
  - ? What do you want to name your add-in? _[AWS Add-in]_
  - ? Which Office client application would you like to support? _[Excel]_
  
**Step 2:**
- Go the directory where your project was created:

         cd C:\Users\priva\AWS Add-in

- Start the local web server and sideload the add-in:

         npm start

- Open the project in VS Code:

         code .
         
## Reference
- [AWS SDK for JavaScript](https://docs.aws.amazon.com/AWSJavaScriptSDK/latest/)

## FAQ
- Q："We can't open this add-in from localhost" when loading an Office add-in
  - A-1: one reason is [reference](https://docs.microsoft.com/en-us/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)
  - A-2: one reason is that you should start local web server. **npm start**
  
