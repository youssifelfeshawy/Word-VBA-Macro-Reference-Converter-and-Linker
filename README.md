# Word VBA Macros: Reference Converter and Linker

This repository contains VBA macros for Microsoft Word that automate the process of converting URLs and DOIs to clickable hyperlinks and linking in-text citations to their corresponding references in the bibliography.

The repository includes three files:
- `ConvertURLsAndDOIsToHyperlinks.bas`: The original macro for converting URLs and DOIs to hyperlinks.
- `LinkCitationsToURLs.bas`: The original macro for linking citations to URLs/DOIs in the bibliography.
- `ConvertAndLinkReferences.bas`: The merged macro that combines the functionality of the two above into a single subroutine.

## Features

- Converts plain-text URLs (http/https) to hyperlinks.
- Converts DOI strings (e.g., "doi: 10.1234/example") to hyperlinks pointing to https://doi.org/.
- Links in-text citations (e.g., [1], [18]) to the URL or DOI in the matching bibliography entry.
- Supports processing the entire document or selected text only.
- Handles multiple occurrences and avoids duplicating existing hyperlinks.

## Requirements

- Microsoft Word (tested on Word for Windows; may work on Mac with adjustments).
- VBA enabled in Word (go to File > Options > Trust Center > Trust Center Settings > Macro Settings, and enable macros).

## Installation

1. Open Microsoft Word.
2. Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.
3. In the VBA editor, go to File > Import File and select the desired .bas file(s) from this repository to import the module(s).
4. Save the module(s) (e.g., as part of your Normal.dotm template for global access).
5. Close the VBA editor.

You can import all three files if you want access to the individual macros and the merged version, or just the merged one for combined functionality.

## Usage

1. Open your Word document containing URLs, DOIs, and citations.
2. If you want to process only a section (e.g., the bibliography), select that text first.
3. Go to Developer > Macros (if the Developer tab isn't visible, enable it in File > Options > Customize Ribbon).
4. Select the desired macro (e.g., `ConvertAndLinkReferences` for the merged version) and click Run.
5. A dialog will ask if you want to process selected text (Yes) or the entire document (No).
6. The macro will run, converting URLs/DOIs and linking citations (for the linking macros).
7. A completion message will show the number of conversions and links made.

### Document Structure Assumptions

- Citations are in the format `[number]` (e.g., [1], [18]).
- Bibliography entries start with `[number]\t` (e.g., [1] followed by a tab).
- The bibliography section starts with a heading "Bibliography".
- URLs are plain http/https links.
- DOIs are in the format "doi: 10.xxxx/xxxxx".

If your document uses a different format, you may need to modify the regex patterns in the code.

## Limitations

- The macros use regular expressions, which may need tweaking for custom formats.
- They loop to handle multiple matches but may be slow on very large documents.
- No error handling for invalid DOIs or URLs; assumes well-formatted input.
- Do not handle citations with ranges (e.g., [1-3]) or non-numeric citations.

## Contributing

Feel free to fork this repository and submit pull requests for improvements, such as better regex patterns or additional features.
