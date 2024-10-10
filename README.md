# My VBA Notebook
A collection of VBA tips - easy to follow, with snippets to copy & paste

<!-- 
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
-->

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

### Collections
It is better to use a dictionary rather than a collection, for the following reasons:
- Performance.
- Richer functionalities.
- Everything you can do with a collection, you can do with a dictionary as well.

Reference: https://www.experts-exchange.com/articles/3391/Using-the-Dictionary-Class-in-VBA.html

### Dictionaries
```VB
Option Explicit

' Add reference: Microsoft Scripting Runtime
Public Sub DictionaryTest()
    Dim oDict As Scripting.Dictionary ' Early binding
    Set oDict = New Scripting.Dictionary

    oDict("Apple") = 5
    oDict("Orange") = 50
    oDict("Peach") = 44
    oDict("Banana") = 47
    oDict("Plum") = 48
    oDict.Add Key:="Pear", Item:="22"
    Call oDict.Add("Strawberry", 11)

    Debug.Print ("There are " & oDict.Count & " items")
    oDict.Remove "Strawberry"
    Debug.Print ("There are " & oDict.Count & " items")

    ' Checks if an item exists by the key
    If Not oDict.Exists("Grapes") Then
        Debug.Print ("This dictionary does not contain grapes")
    End If

    Set oDict = Nothing
End Sub
```

- Adding the same key more than once, will result in an error.
- If you use the Item property to attempt to set an item for a non-existent key, the Dictionary will implicitly add that
item along with the indicated key.
- Similarly, if you attempt to retrieve an item associated with a non-existent key, the Dictionary will add a blank item,
associated with that key.
- CompareMode is used to compare the keys: Binary vs Text Compare.

### Traversing the Dictionary
```VB
Dim key As Variant

For Each key In oDict.Keys
    Debug.Print key & " - " & oDict(key)
Next
```

### Removing a key
The Remove method removes the item associated with the specified key from the Dictionary, as well as that key.
```VB
MyDictionary.Remove "SomeKey"
```

### Clear the dictionary
```VB
MyDictionary.RemoveAll
```

## Boosting Performance
### Speeding the read and write process from cells
- Read data in ranges.
- Turn screen updating off
- Turn calculation off
- Read and write the range at once

```VB
Sub Datechange()
    On Error GoTo error_handler
    
    Dim initialMode As Long
    
    initialMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Dim data As Variant
    Dim i As Long

    'copy range to an array
    data = Range("D2:D" & Range("D" & Rows.Count).End(xlUp).Row)

    For i = LBound(data, 1) To UBound(data, 1)
        If IsDate(data(i, 1)) Then data(i, 1) = CDate(data(i, 1))
    Next i

    'copy array back to range
    Range("D2:D" & Range("D" & Rows.Count).End(xlUp).Row) = data

exit_door:
    Application.ScreenUpdating = True Application.Calculation = initialMode
    Exit Sub

error_handler:
    'if there is an error, let the user know
    MsgBox "Error encountered on line " & i + 1 & ": " & Err.Description
    Resume exit_door 'don't forget the exit door to restore the calculation mode
End Sub
```

### Clearing Ranges
When clearing cells in Excel and we already know which range needs to be cleared, it is much faster to use the .Clear method on the predefined range, rather than clearing cell by cell.
```VB
Thisworkbook.Sheets(1).Range("A1:J999").Clear
```

### Calculating elapsed time in seconds
```VB
Private Sub Process()
    Dim tickStart As Date: tickStart = Now()
    Dim tickEnd As Date
    
    ' Processing goes here
    tickEnd = Now()
    
    MsgBox DateDiff("s", tickStart, tickEnd)
End Sub
```

### Mergesort
```VB
Option Explicit

Const MaxN As Long = 100000
Dim a(1 To MaxN) As Long
Dim tmp(1 To MaxN) As Long

Private Sub Mergesort(ByVal l As Long, ByVal r As Long)
    If (r > l) Then
        Dim mid As Long: mid = (r + l) \ 2
        Call Mergesort(l, mid)
        Call Mergesort(mid + 1, r)
    
        Dim i As Long, j As Long, k As Long
        i = l
        j = mid + 1
        k = 1
        
        Do While (i <= mid And j <= r)
            If (a(i) > a(j)) Then
                tmp(k) = a(j)
                j = j + 1
            Else
                tmp(k) = a(i)
                i = i + 1
            End If
            
            k = k + 1
        Loop
        
        Do While (i <= mid)
            tmp(k) = a(i)
            i = i + 1
            k = k + 1
        Loop
        
        Do While (j <= r)
            tmp(k) = a(j)
            j = j + 1
            k = k + 1
        Loop
        
        For i = 1 To r - l + 1
            a(l + i - 1) = tmp(i)
        Next i
    End If
End Sub

Public Sub Test()
    Dim i As Long
    Dim tickStart As Date: tickStart = Now()
    Dim tickEnd As Date
    
    For i = 1 To MaxN
        a(i) = Rnd * MaxN
    Next i
    
    Call Mergesort(1, MaxN)
    
    For i = 2 To MaxN
        Debug.Assert a(i) >= a(i - 1)
    Next i
    
    tickEnd = Now()
    Debug.Print "Time taken: " & DateDiff("s", tickStart, tickEnd)
End Sub
```

## File Handling
### Selecting a file via the File Dialog
The File Dialog is used to select files by browsing the computer. It also allows multiselect, give the possibility to add filters so that we have a choice of which kind of files can be selected, etc...

```VB
Sub UseFileDialogOpen()
    Dim lngCount As Long
    
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        ' .AllowMultiSelect = True
        .AllowMultiSelect = False
        .Show
        .Filters.Add "Txt", "*.txt"
        
        If .SelectedItems.Count = 1 Then
            ThisWorkbook.Sheets("Instructions").Cells(15, 6).Value = .SelectedItems(1)
        Else
            ThisWorkbook.Sheets("Instructions").Range("G15:G15").Clear
        End If
        ' Display paths of each file selected
        ' For lngCount = 1 To .SelectedItems.Count
        ' MsgBox .SelectedItems(lngCount)
        ' Next lngCount
    End With
End Sub
```

### Reading from an input file
```VB
Public Sub ReadFile()
    Dim myfile As String: myfile = "..."
    Dim textline As String
    Dim linecount As Long: linecount = 0
    
    Close #1
    Open myfile For Input As #1
    
    Do Until EOF(1)
        Line Input #1, textline
        linecount = linecount + 1
    Loop
    
    Debug.Print linecount
    Close #1
End Sub
```

### Writing to an output file
```VB
Public Sub WriteToFile()
    Dim myfile As String: myfile = "c:\users\x76544\try.txt"
    Close #1

    Open myfile For Output As #1
    Print #1, "This is a test" ' Outputs to file without double quotes
    Write #1, "This is a test" ' Outputs to file with double quotes
    
    Close #1
End Sub
```

### Getting a file extension
```VB
    Set oFs = New FileSystemObject
    .
    .
    For Each oFile In currentFolder.Files
    .
    .
    Debug.Print oFs.GetExtensionName(oFile.path)
Next
```

### Recursively get a list of files
Firstly, we should add a reference to the DLL “Microsoft Scripting Runtime”.
This DLL exposes the “FileSystemObject” class, which will be used for traversing the folders recursively.
The following example traverses a folder, picks up all the .cpp files and count the number of lines each file contains.

```VB
Sub CountLines(oFile As File)
    Dim oTextStream As TextStream
    Dim lineCount As Long: lineCount = 0

    Set oTextStream = oFile.OpenAsTextStream(ForReading)

    Do While Not (oTextStream.AtEndOfStream)
        oTextStream.ReadLine
        lineCount = lineCount + 1
    Loop

    fileNum = fileNum + 1
End Sub

Sub Traverse(currentFolder As Folder)
    Dim oFile As File
    Dim oFolder As Folder
    
    ' Gets the list of .cpp files in the current folder
    For Each oFile In currentFolder.Files
        If (oFile.Type = "CPP File") Then
        ' Code goes here...
        End If
    Next
    
    ' Recurse in each folder
    For Each oFolder In currentFolder.SubFolders
        Call Traverse(oFolder)
    Next
End Sub

Public Sub Test()
    Dim oFS As Scripting.ileSystemObject
    Set oFS = New FileSystemObject
    
    Call Traverse(oFS.GetFolder("..."))
    
    Set oFS = Nothing
End Sub
```

### Copying files & folders
```VB
Dim ofs As New FileSystemObject
ofs.CopyFile "Source File", "Destination File"

Set ofs = Nothing
```

The FileSystemObject also exposes other interesting methods like to copy folders, create folders etc.

### Connection to Database
Connecting to the local MS Access database in VBA
Reference: https://msdn.microsoft.com/en-us/library/office/ff835631.aspx

```VB
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strSQL As String

' Use the current db
Set db = CurrentDb

' Build the sql query
strSQL = "SELECT * FROM Person"

' Execute the query
Set rs = db.OpenRecordset(strSQL)

' Traversing the dataset result
Do While Not rs.EOF
    Debug.Print rs!Id & " " & rs!firstname & " " & rs!familyname
    rs.MoveNext
Loop

Debug.Print rs.RecordCount

' Cleaning up
rs.Close
db.Close
```

## Dealing with MS Office & PDF files
### Microsoft Excel
#### Creating an Excel File
```VB
Dim oXlsxApplication As Excel.Application
Dim oXlsxWorkbook As Excel.Workbook
Dim oXlsxWorksheet As Excel.Worksheet

Set oXlsxApplication = New Excel.Application
Set oXlsxWorkbook = oXlsxApplication.Workbooks.Add
Set oXlsxWorksheet = oXlsxWorkbook.Sheets.Add

' Code goes here
' ...

Set oXlsxApplication = Nothing
Set oXlsxWorkbook = Nothing
Set oXlsxWorksheet = Nothing
```

## Microsoft Word
### Creating a Word Document
```VB
Dim oWordApplication As Word.Application
Dim oWordDocument As Word.Document

Set oWordApplication = New Word.Application
Set oWordDocument = oWordApplication.Documents.Add

With oWordDocument
    .Content.InsertAfter "This is a test"
End With

oWordApplication.Visible = True
```

## Outlook
### References
To use the outlook object, make sure the “Microsoft Outlook 15.0 Object Library” is added as reference.

![Microsoft outlook Object Library reference](/assets/images/vba_references.png)

### Sending emails via Outlook
```VB
Dim locObjOutlook As Outlook.Application
Dim locObjOutlookItem As Outlook.MailItem
Dim locObjOutlookItemCopy As Outlook.MailItem
Dim htmlBody As String: htmlBody = ""

Set locObjOutlook = New Outlook.Application
Set locObjOutlookItem = locObjOutlook.CreateItem(olMailItem)

locObjOutlookItem.BodyFormat = olFormatHTML
htmlBody = htmlBody & "<html>"
htmlBody = htmlBody & " <head>"
.
.
htmlBody = htmlBody & " </head>"
htmlBody = htmlBody & " <body>"
.
.
htmlBody = htmlBody & " </body>"
htmlBody = htmlBody & "</html>"

locObjOutlookItem.htmlBody = htmlBody
locObjOutlookItem.Display ' displays the email first
Set locObjOutlook = Nothing
```

## Creating a PDF File
We can simulate the creation of a pdf file by first creating an office file and then using the “Save” command to save it as
a pdf.
For saving a file under the pdf format, we use file format = 17.

```VB
Dim oWordApplication As Word.Application
Dim oWordDocument As Word.Document

Set oWordApplication = New Word.Application
Set oWordDocument = oWordApplication.Documents.Add

With oWordDocument
    .Content.InsertAfter "This is a test"
    .SaveAs2 "C:\Users\x76544\" & "myDoc.pdf", FileFormat:=17
End With

oWordDocument.Close

Set oWordApplication = Nothing
```