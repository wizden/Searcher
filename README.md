## Description
This utility is intended to help in search/grep for multiple content in multiple locations for multiple file types.

## Main features
* Search multiple directories for multiple content based on multiple file extensions.
* Allows searching in Word (docx), Excel (xlsx), Powerpoint (pptx) and PDF (pdf) documents.
* Allow searching in archives - .zip, .7z, .rar, .tar and .gz
* Page numbers in Word (docx) may not be exact using OpenXML(see http://officeopenxml.com/WPsection.php)
* Search using regular expressions.
* Shows list of files that were searched.
* Exclude files/directories from being searched in the current session or always.
* Open text files in preferred editor.
* Navigate to file location in explorer.
* Quickly view file content in a popup.
* Advanced configuration allows changing backcolour and displaying search execution times.
* Customise the UI based on settings in the preferences file (if set)
* No installer. Setup / installation not required.

Searcher  is free software distributed under the GNU GPL. Ensure that the distribution package contains the following files:
* Searcher.exe  - File content search utility
* COPYING.txt   - License information
* README.txt    - README file

Special thanks to the creators of the following libraries that is used by Searcher:
   - Costura Fody  - combine libraries into executable
   - iTextSharp    - reading pdf files.
   - OpenXml       - reading files saved using openxml format.
- SharpCompress - library to deal with many compression types and formats.

## Additional Links
[Using Searcher](https://github.com/wizden/Searcher/wiki/Using-Searcher)

![Build status](https://ci.appveyor.com/api/projects/status/github/wizden/Searcher?svg=true)
