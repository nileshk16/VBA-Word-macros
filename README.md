# VBA-Word-macros
# VBA Script for Opening URLs or Searching on Google

This script consists of two main procedures written in VBA (Visual Basic for Applications). It allows users to open a specified URL in Google Chrome or perform a Google search with selected text from a Word document. Below is the description of each part of the code:

## Sub OpenBrowser(strAddress As String)

This procedure is responsible for opening Google Chrome with the specified URL.

- **Parameters:**
  - `strAddress`: The URL or search query to be opened in Google Chrome.

- **Functionality:**
  - Defines the path to the Chrome executable.
  - Uses the `Shell` function to open Chrome with the given URL.

### Steps:
1. Define the path to the Chrome executable.
2. Use the `Shell` function to open a new Chrome process with the specified URL.

## Sub SearchOnGoogle()

This procedure performs a Google search or opens a URL based on the selected text in a Word document.

- **Functionality:**
  - Checks if text is selected in the Word document.
  - If text is selected, it trims any extra spaces and copies the text.
  - Determines if the selected text is a URL or a search query.
  - If the text is a URL, it opens the URL in Chrome.
  - If the text is not a URL, it performs a Google search with the selected text.

### Steps:
1. Verify if text is selected in the Word document.
   - If no text is selected, display a message box prompting the user to select text.
2. If text is selected, trim extra spaces and copy the text.
3. Check if the selected text is a URL (starting with "https://", "http://", or "www.").
   - If it is a URL, open it in Chrome.
   - If it is not a URL, perform a Google search with the selected text.

## Usage

1. Select the text in the Word document that you want to search or open as a URL.
2. Run the `SearchOnGoogle` procedure.
3. The script will determine if the text is a URL or a search query and open the appropriate page in Google Chrome.

This script is useful for quickly searching the web or opening links directly from a Word document, streamlining the workflow and enhancing productivity.
