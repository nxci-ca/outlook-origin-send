# Outlook add-in to check external recipients

This Outlook add-in enhances the external recipients check experience, by adding an extra layer of security to avoid leaking confidential information to people outside your organization.

![image (4)](https://user-images.githubusercontent.com/1230332/231283775-efd166c8-4b60-4082-bce6-e11c58c8b07e.png)

## Prerequisites

You can reference [the official documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview) to identify the prerequisites for your development environment.

## Setup

1. Clone the repository
2. Open the project's folder in a terminal and run:

   ```bash
   npm install
   ```
3. In the **src/launchevent/launchevent.js** file, change the value of the `customerDomain` property to match your organizational's domain (e.g. @contoso.com).
4. Run the following command to compile the project:
   ```bash
   npm run build
   ```
5. open the dist folder
   ```bash
   cd dist
   ```
6. run the following command to execute the project:
   ```bash
   npm start
   ```
7. Compose a new e-mail, add at least one e-mail address which doesn't belong to the domain you have specified in Step 3 and press Send.

## Deploy
To deploy the application, it should be installed from admin Managed addons using the following url:
https://aka.ms/olksideload
## Learn more
This project is described in details on the [Modern Work App Consult blog](https://techcommunity.microsoft.com/t5/modern-work-app-consult-blog/a-better-way-to-identify-external-users-in-an-outlook-mail/ba-p/3793131).
