# excel-copilot

An Excel task pane add-in that uses Claude AI to autocomplete cells — formulas, values, or text — based on the current spreadsheet context.

## How it works

Type in a cell or enter a prompt manually in the task pane. The add-in reads the active worksheet, builds a context-aware prompt, and calls the Claude API to suggest a completion. Accept to write it to the selected cell, or dismiss.

## Quickstart

**Prerequisites:** Node.js, Microsoft Excel (desktop), an [Anthropic API key](https://console.anthropic.com/).

```bash
# 1. Install dependencies
npm install

# 2. Set your API key in src/taskpane/taskpane.ts
#    Replace "YOUR_API_KEY" with your Anthropic API key

# 3. Start the dev server (HTTPS on port 3000)
npm run dev-server

# 4. Sideload the add-in in Excel
npm run start
```

The add-in button appears in the **Home** tab ribbon. Click **Show Task Pane** to open it.

To stop:

```bash
npm run stop
```

## Build

```bash
npm run build        # production
npm run build:dev    # development
npm run watch        # watch mode
```

Before a production build, update the deployment URL in [webpack.config.js](webpack.config.js):

```js
const urlProd = "https://your-domain.com/";
```

## License

MIT License

Copyright (c) 2024

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
