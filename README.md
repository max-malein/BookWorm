# BookWorm
Bookworm is a Grasshopper plugin to read and write data from Google Spreadsheets.

## Installation
The easiest way to install is to use a Package Manager. Go to Tools/PackageManager, then search for bookworm and click install.

## Usage
You can download example files from here https://github.com/max-malein/BookWorm/tree/master/GH_tests

When first launched you will be asked to authorize the app with your google account.
The app is a work in progess, so you will get some warnings from Google that it hasn't been reviewed by Google and bla-bla-bla... You jusn need to confirm that you're OK with it.

After authorization your personal token will be saved to C:\Users\\*UserName*\AppData\Roaming\McNeel\Rhinoceros\packages\7.0\bookworm\\*version_number*\token.json

You will be able to read and edit only spreadsheets you have permission to do so with your google account. 
Please consider making your spreadsheet public if you plan to share your grasshopper definition with others.

In case you need to change your google account for BookWorm, simply delete the token.json folder.
You will be asked to login again next time the grasshopper solution is recomputed.

## Development
If you found a bug or just have some ideas how to make BookWorm better please drop me a note at max@malein.com

This is an open source project and pull requests are always welcome.

-- Max Malein
