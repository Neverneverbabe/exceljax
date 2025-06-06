# ExcelJax

This project is an Excel add‑in. The build output is committed to the `docs/` directory so it can be hosted on GitHub Pages.

## Setup

1. Install dependencies:
   ```bash
   npm install
   ```
2. Open `webpack.config.js` and change the `urlProd` constant to the URL of your GitHub Pages site. It should include the trailing slash, for example:
   ```js
   const urlProd = "https://<username>.github.io/<repo>/";
   ```
3. Build the project. You can use either command:
   ```bash
   npm run build       # generic production build
   npm run build:ghpages # build using the GitHub Pages configuration
   ```
   The compiled files will be written to the `docs/` directory (or `dist/` depending on the script).
4. Commit and push the generated files to GitHub. Enable GitHub Pages for the repository and set the source to the `docs` folder on the `main` branch (or the branch you use).
5. After deploying, sideload the `docs/manifest.xml` file in Excel. In Excel, choose **Insert → My Add-ins → Shared Folder → Upload My Add-in** and select the manifest file from the `docs` directory.

## Local testing

For offline or local testing you need an HTTPS server. You can start one with:

```bash
npm start                 # uses office-addin-debugging to host locally
# or
npx http-server docs -S   # serve the built files from the docs folder
```

Office requires add-ins to run from HTTPS, so make sure the local server uses HTTPS certificates.

## Task script

Use `scripts/tasks.js` to run common project tasks from the command line:

```bash
node scripts/tasks.js install       # run npm install
node scripts/tasks.js build         # build for production
node scripts/tasks.js buildGhPages  # build for GitHub Pages
node scripts/tasks.js proxy         # start the local LLM proxy
node scripts/tasks.js start         # launch the add-in locally
```

## Testing

Run lint checks using:

```bash
npm test
```

To manually verify the add-in:

1. Build the project and write the output to the `docs/` directory:
   ```bash
   npm run build:ghpages
   ```
2. In Excel, sideload `docs/manifest.xml` by choosing **Insert → My Add-ins → Shared Folder → Upload My Add-in**.
3. Interact with the add-in to ensure the commands work as expected. During development you can serve the `docs` folder over HTTPS:
   ```bash
   npm start                 # uses office-addin-debugging
   # or
   npx http-server docs -S   # serve the built files
   ```

