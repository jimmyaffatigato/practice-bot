# Practice Bot

Practice Bot is Google Apps Script Project for generating practice charts.

It uses the Google Classroom API v1 and the Google Sheets API v4 to automate the creation and distribution of practice charts.

## Environment

This project's source code is written in Typescript. All necessary declaration files are provided by the npm package `google-apps-script` as a dev dependency.

It requires OAuth2 authentication to push code to a Google account. It also requires the Sheets and Classroom services to be enabled by the user in their Apps Script configuration.

This project was developed by Jimmy Affatigato for use in the City School District of Albany. Reuse and modification are granted by the MIT License.

## Build

`clasp push`

[Clasp](https://github.com/google/clasp) is configured to upload only the entry point `main.ts` and the manifest file `appsccript.json`. All code must be included in `main.ts` to be used in the project. Clasp and Google Apps Script do not support ES6 modules.

Clasp automatically transpiles `.ts` files to V8-compatible ES6 for Google Apps. It is not necessary to run `tsc` prior to pushing code.
