# My VBA Notebook
A collection of VBA tips - easy to follow, with snippets to copy & paste

# Contents
[Coding Best Practices](#coding-best-practices)
[Force variable declarations](#force-variable-declarations)
Error Handling
Null values
DebugAssert
The “Not Responding” problem
Getting the containing folder of the tool
Generating random numbers
Object oriented coding style
Class Description
Using an instantiated class
Data Structures
Static array
Dynamic array
Array keyword
Create an array using the split keyword
Looping through an array
Check if an array is allocated
Collections
Dictionaries
Traversing the Dictionary
Removing a key
Clear the dictionary
Boosting Performance
Speeding the read and write process from cells
Clearing Ranges
Calculating elapsed time in seconds
Binary search the last filled row / column
Sorting: Mergesort
File Handling
Selecting a file via the File Dialog
Reading from an input file
Writing to an output file
Getting a file extension
Recursively get a list of files
Copying files & folders
Connection to Database
Connecting to the local MS Access database in VBA
Dealing with MS Office & PDF files
Microsoft Excel
Creating an Excel File
Microsoft Word
Creating a Word Document
Outlook
References
Sending emails via Outlook
Creating a PDF File
Internet Explorer Automation
Required References
Windows API
Loading a new Internet Explorer Window and navigate to wwwgooglecom
Check if any of the opened Internet Explorer windows is already on a specific page
Document object
Searching for an HTML Element by its ID
Common HTML objects
Searching for HTML elements by its type
Generic Robot Class
Waiting in the application
Force the robot to click on “Yes” on a confirmation window
ActiveX controls
Formatting Data
Padding with leading zeros
MS DOS
Getting help for a particular command
Accessing a folder on the network
Executing a command on selected files
Useful techniques
Hiding sheet tabs
Hiding Row numbers and Column numbers

## Coding Best Practices
### Force Variable declarations
```VB
Option Explicit ' Always include this at the top each source file
```

### Error Handling
```VB
Public Function Foo(...) As Boolean
    Const strPROC_NAME As String = "Foo"

On Error GoTo Error_handler
    ' My code goes here
    ' If everything goes on perfectly, exit the function smoothly
    Foo = True
    Exit Function

Error_handler:
    MsgBox "An error occured ...: " & Err.Description
    Foo = False
    Exit Function
End Function
```

### Null values
To check if a value is null, use the IsNull(..) function.

### Debug.Assert
Assertions are used in development to check your code as it runs. An Assertion is a statement that evaluates to true or
false. If it evaluates to false then the code stops at that line. This is useful as it stops you close to the cause of the error.
```VB
Debug.Assert 1 = 2
```

### The “Not Responding” problem
Reference: https://support.microsoft.com/en-us/kb/118468

When a time consuming program runs, most of the time, Excel will fall in a “Not Responding” state, although the
program continues to run in the background. In such situation, we would like to have a kind of progress feedback on the
screen so that we are sure the program is not stuck in an infinite loop. In such case, use the command:
```VB
DoEvents
```

### Getting the containing folder of the tool
We need to often output files to a folder at the same level of the tool. It is better NOT to hardcode that path in the
code. Instead, use the following command to get the path of the Workbook.
```VB
ThisWorkbook.Path & "\MyOutputFolder\" & OutputFilename & ".txt"
```

### Generating random numbers
Use the function from the Worksheet object to generate random numbers.
```VB
WorksheetFunction.RandBetween(1, 10000)
```