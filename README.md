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

## Object oriented coding style
### Class Description
```VB
'
' Class : Robot
' Description : Generic class for Robot
'
Option Explicit

Private Sub class_initialize()
    ' Constructor
    Debug.Print "Robot initialized"
End Sub

Private Sub class_terminate()
    ' Destructor
    Debug.Print "Robot destroyed"
End Sub
```

### Using an instantiated class
```VB
Option Explicit

Public Sub GO()
    Dim oRobot As Robot
    
    ' Launch Robot for the simulation
    Set oRobot = New Robot
    
    ' Release memory
    Set oRobot = Nothing
End Sub
```

## Data Structures
### Static Array
```VB
Public Sub DecArrayStatic()
    Dim arrMarks1(0 To 3) As Long ' Create array with locations 0,1,2,3
    Dim arrMarks2(3) As Long ' Defaults as 0 to 3 i.e. locations 0,1,2,3
    Dim arrMarks1(1 To 5) As Long ' Create array with locations 1,2,3,4,5
    Dim arrMarks3(2 To 4) As Long ' Create array with locations 2,3,4
End Sub
```

### Dynamic array
```VB
Public Sub DecArrayDynamic()
    Dim arrMarks() As Long ' Declare dynamic array
    ReDim arrMarks(0 To 5) ' Set the size of the array when you are ready
End Sub
```

### Array keyword
```VB
Public Sub DeclareArray()
    ' To create and "Array", use the Variant keyword
    Dim arr1 As Variant
    arr1 = Array("Orange", "Peach", "Pear")

    Dim arr2 As Variant
    arr2 = Array(5, 6, 7, 8, 12)
End Sub
```

### Create an array using the split keyword
```VB
public Sub DeclareArrayUsingSplit()
    Dim s As String
    s = "Red,Yellow,Green,Blue"

    Dim arr() As String
    arr = Split(s, ",")
End Sub
```

### Looping through an array
```VB
Public Sub ArrayLoops()
    Dim arrMarks(0 To 5) As Long
    Dim i As Long
    
    For i = LBound(arrMarks) To UBound(arrMarks)
        arrMarks(i) = 5 * Rnd ' Fill the array with random numbers
    Next i
End Sub
```

The functions LBound and UBound are very useful. Using them means our loops will work correctly with any array size.
The real benefit is that if the size of the array changes we do not have to change the code for printing the values. A loop
will work for an array of any size as long as you use these functions.

```VB
For Each mark In arrMarks
    mark = 5 * Rnd ' Will not change the array value
Next mark
```

### Check if an array is allocated
Sometimes, an array is declared without dimensions and grows dynamically with the ReDim keyword. That array may
stay without being re-dimensioned. Using the LBound(..) or UBound(..) function on that array will throw the “Subscript
out of range error”. A solution is to use the following snippet before using the LBound or UBound functions.

```VB
Dim myArray() As String 'Declare array without dimensions

If (Not Not myArray) = 0 Then 'Means it is not allocated
.
.
Else
.
.
End if
```
