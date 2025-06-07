# ExcelJax

ExcelJax is an experimental Excel Office Add-in powered by LM Studio.  
It allows you to use local LLMs running in LM Studio for data analysis, formula assistance, summarization, and more, directly in Excel.

## Features

- Dynamically loads available models from LM Studio (`http://localhost:1234/v1/models`)
- AI chat, formula help, insights, and data cleanup tools in the Excel Task Pane
- Deployed and updated automatically via GitHub Actions to GitHub Pages

## Quick Start

1. **Install dependencies:**  
   `npm ci`

2. **Development build:**
   `npm run build:dev`

3. **Deploy:**  
   Just push to `main` — GitHub Actions builds and publishes automatically!

4. **Update manifest:**  
   Download `manifest.xml` from your repo and [sideload it in Excel](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

5. **Start LM Studio:**  
   Make sure LM Studio is running and listening on `http://localhost:1234`.

## Available Scripts

- `npm run build` – Production build with Webpack.
- `npm run build:dev` – Development build.
- `npm run build:ghpages` – Build using the GitHub Pages configuration.
- `npm run dev-server` – Run a development server with live reload.
- `npm run watch` – Watch source files and rebuild on changes.
- `npm run start` – Launch the add-in for debugging.
- `npm run stop` – Stop the debugging session.
- `npm run lint` – Check code style with ESLint.
- `npm run lint:fix` – Fix fixable linting issues.
- `npm run proxy` – Start the LM Studio proxy server.

## Project Structure

- `/src` — source files (`taskpane.html`, `taskpane.js`)
- `/dist` — production build output deployed from CI
- `/docs` — static site generated during the build
- `/dist-ghpages` — GitHub Pages build output
- `.github/workflows/deploy.yml` — CI/CD for automatic deployment

`dist/`, `docs/`, and `dist-ghpages/` are generated automatically by the workflow
and are intentionally ignored in Git.

## License

MIT
