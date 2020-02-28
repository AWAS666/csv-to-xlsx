# csv-to-xlsx
Creating Excel sheets from csv files

Uses xlnt for writing xlsx files, you can get that from [vcpkg](https://github.com/Microsoft/vcpkg)

Can be called from command line, or just drag and drop your files onto the finished exe



Known issues:
- ANSI to UTF-8 needs some proper coding, currently only handles commonly used german vocals
- everything is written as a string into the xlsx file
