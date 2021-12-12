# Issues Export Import

This is an example that shows how you can export the issues, from the Issue Resolution Service, to excel using a 3rd party library.

It also allows a structured format spreadsheet (preferably using the export tool) to be imported into the system creating new issues or updating existing ones.

It contains no retry logic and minimal logging/error handling and uses a 3rd party library.

This sample is provided as is without warranty of any kind and it use is at your own risk. 

## Environment Variables

Prior to running the app, update the AuthorizationClient.ts file with your app id and optional extra scopes. Note minimal scopes shown below:

```
    const scope = "itwinjs email openid profile organization issues:modify issues:read projects:read urlps-third-party users:read";
    const clientId = "spa-xxxxxxxxxxxx";
```

You can also replace the OIDC client data in this file with your own if you'd prefer.

## Available Scripts

In the project directory, you can run:

### `npm start`

Runs the app in the development mode.\
Open [http://localhost:3000](http://localhost:3000) to view it in the browser.

The page will reload if you make edits.\
You will also see any lint errors in the console.

### `npm test`

Launches the test runner in the interactive watch mode.\
See the section about [running tests](https://facebook.github.io/create-react-app/docs/running-tests) for more information.

### `npm run build`
Builds the app for production to the `build` folder.\
It correctly bundles React in production mode and optimizes the build for the best performance.

The build is minified and the filenames include the hashes.\
Your app is ready to be deployed!

See the section about [deployment](https://facebook.github.io/create-react-app/docs/deployment) for more information.

### `npm run eject`

**Note: this is a one-way operation. Once you `eject`, you can’t go back!**

If you aren’t satisfied with the build tool and configuration choices, you can `eject` at any time. This command will remove the single build dependency from your project.

Instead, it will copy all the configuration files and the transitive dependencies (webpack, Babel, ESLint, etc) right into your project so you have full control over them. All of the commands except `eject` will still work, but they will point to the copied scripts so you can tweak them. At this point you’re on your own.

You don’t have to ever use `eject`. The curated feature set is suitable for small and middle deployments, and you shouldn’t feel obligated to use this feature. However we understand that this tool wouldn’t be useful if you couldn’t customize it when you are ready for it.

## Learn More

You can learn more in the [Create React App documentation](https://facebook.github.io/create-react-app/docs/getting-started).

To learn React, check out the [React documentation](https://reactjs.org/).
