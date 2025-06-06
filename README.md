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
   `npm run build`

3. **Deploy:**  
   Just push to `main` — GitHub Actions builds and publishes automatically!

4. **Update manifest:**  
   Download `manifest.xml` from your repo and [sideload it in Excel](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

5. **Start LM Studio:**  
   Make sure LM Studio is running and listening on `http://localhost:1234`.

## Project Structure

- `/src` — source files (`taskpane.html`, `taskpane.js`)
- `/dist` — built output, deployed to GitHub Pages (`main` branch → `gh-pages` branch)
- `.github/workflows/deploy.yml` — CI/CD for automatic deployment

## License

MIT