# tweetsearch2xlsx
Parse a twitter search (saved as html file) and return a MS Excel file (xlsx). Note that this is currently an experimental project.

Completely unoptimized and thus can be resource-intensive.

It is also very inflexible, as it currently does not treat tweets as individual units but instead composes the final file out of a list of authors, messages, dates, etc..

##Usage
```
$ ./tweetsearch2xlsx.py [-o outputfile.xlsx] inputfile.html
```

If output is not specified, it will assume the same filename as input is desired with the correct extension.

##Installation
This python3 script relies on third-party libraries lxml and xlsxwriter.

You can find an unnofficial binary version of lxml for Windows platforms at http://www.lfd.uci.edu/~gohlke/pythonlibs/#lxml
