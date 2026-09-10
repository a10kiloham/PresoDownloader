# DocSend & Sharepoint Slide Deck Downloader
A Chrome extension to download DocSend and Sharepoint slides as a PDF directly from a logged in browser.

## Installation from source

1. Clone this repository to your machine as a .zip file.
2. Navigate to `chrome://extensions/` on Google Chrome.
3. Toggle Developer Mode on the top right if not toggled already.
4. Drag and drop the zip file to the exensions window.

## Usage

Navigate to any DocSend slide deck page or Sharepoint slide deck page and click on the extension icon on the top right of Google Chrome to generate a PDF of the slide deck. A small status panel appears in the bottom-right corner of the page while the download runs. If the file cannot be fetched, the panel shows "File Unreachable" for five seconds.

* **DocSend**: open the slide deck page (enter your email if the deck asks for it). Each slide image is fetched and assembled into a PDF. Slide image links are requested one at a time right before use, so long decks no longer fail with expired links.
* **SharePoint / OneDrive**: open the file so it is showing in the SharePoint viewer (the URL contains `id=/sites/.../file.pdf`). The extension fetches the original PDF through the same media service the viewer uses, so the download is the real file, not a screenshot. This works for files where SharePoint hides the Download button but you can still view them.
  * PDFs are saved as-is.
  * Word, PowerPoint and Excel files are converted to PDF by SharePoint's media service (experimental).

## How the SharePoint download works

1. The extension reads `_spPageContextInfo` from the page to find the site and library, and resolves the open file to its list item.
2. It calls `RenderListDataAsStream` with the same render options as the built-in viewer. The response carries a short-lived media token, the regional media host and a URL template.
3. It fills the template with `method=passthrough` and fetches the file from the media host with the token in the `X-SPOPacToken` header. The response is the original file bytes, which are saved with a normal browser download.
