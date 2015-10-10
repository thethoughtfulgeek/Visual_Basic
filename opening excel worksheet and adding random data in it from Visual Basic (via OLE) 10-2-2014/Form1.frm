VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
End
End Sub

Private Sub Command1_Click()
 Dim names(5, 2) As String
 Dim oXL As Object, oWB As Object, oSheet As Object
' Here oXL,oWB,... etc are declared as general objects.
' They are later declared as specific excel objects in the program.
' eg. oXl becomes an Excel Application object, oWB becomes an Excel Workbook object
' OSheet an Excel Worksheet object and ORng an excel Range object
' This is known as LATE-BINDING!
' There are two ways to control an Automation server: by using either late binding or early binding.
' Declaring an object variable with the "As object" clause creates a variable that can contain a reference
' to any type of object.
'However, access to the object through that variable is late bound;that is, the binding occurs when your program is run.
' To create an object variable that results in early binding, that is, binding when the program is compiled,
' declare the object variable with a specific class ID.
' for EARLY BINDING the objects should be declared in this manner
' Dim oXL As Excel.Application
' Dim oWB As Excel.Workbook
' Dim oSheet As Excel.Worksheet
' Dim oRng As Excel.Range
    Set oXL = CreateObject("Excel.Application")
' This code starts the application creating the object, in this case a Microsoft excel application
' Once an application object is created, it is referenced to the oXL object variable.
' So now, you can access the properties,objects and methods as well as events of the Excel Application using oXL
' object variable.
    oXL.Visible = True
    Set oWB = oXL.Workbooks.Add
' we are adding workbook to excel application via referencing it through oXL object variable and assigning that workbook
' object to oWB object variable via Set keyword.
    Set oSheet = oWB.ActiveSheet
' activesheet object is assigned to oSheet object variable
    oSheet.Cells(1, 1).Value = "First Name"
'Add table headers going cell by cell
    oSheet.Columns("A:A").ColumnWidth = 10
'formatting column width
    oSheet.Cells(1, 2).Value = "Last Name"
    oSheet.Columns("B:B").ColumnWidth = 10
    oSheet.Cells(1, 3).Value = "Full Name"
    oSheet.Cells(1, 4).Value = "Salary"
'Add table headers going cell by cell
' Format cells by changing fonts and colors now
    oSheet.Range("A1", "D1").Font.Bold = True
    oSheet.Range("A1", "D1").VerticalAlignment = xlVAlignCenter
    oSheet.Range("A1", "D1").Interior.Color = vbRed
    oSheet.Range("A1", "D1").Font.Color = vbYellow
' provide borders to the header columns
    oSheet.Range("A1", "D1").Select
    With Selection
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
' Create an array to set multiple values at once.
    names(0, 0) = "John"
    names(0, 1) = "Smith"
    names(1, 0) = "Tom"
    names(1, 1) = "Brown"
    names(2, 0) = "Sue"
    names(2, 1) = "Thomas"
    names(3, 0) = "Jane"
    names(3, 1) = "Jones"
    names(4, 0) = "Adam"
    names(4, 1) = "Johnson"
' Fill A2:B6 with an array of values (First and Last Names).
    oSheet.Range("A2", "B6").Value = names
' Fill C2:C6 with a relative formula
    oSheet.Range("C2", "C6").Formula = "=A2 & "" "" & B2"
    Columns("C:C").ColumnWidth = 13
' Fill D2:D6 with a formula(=RAND()*100000) and apply format.
    oSheet.Range("D2", "D6").Formula = "=RAND()*100000"
' generates random number upto the limit 100000
    oSheet.Range("D2", "D6").NumberFormat = "[$$-409]#,##0.00"
' $ generates number with string format while # will generate it as a number
      Call DisplayQuarterlySales(oSheet)
' Manipulate a variable number of columns for Quarterly Sales Data
End Sub
Private Sub DisplayQuarterlySales(oWS As Excel.Worksheet)
' We are passing an excel worksheet (the active sheet in which we worked) to a new function DisplayQuarterlySales
' In order to pass the activesheet we use the reference variable oSheet and pass it via DisplayQuarterlySales function
' it is passed to a new object variable oWS
' now we need to declare this object also as a worksheet object which is done by Excel.Worksheet word
    Dim oResizeRange As Excel.Range
' new variable for referencing range object in a different sheet
    Dim oChart As Excel.Chart
' new variable for referencing chart object in a different sheet
    Dim iNumQtrs As Integer
    Dim sMsg As String
    Dim iRet As Integer
' Determine how many quarters to display data for.
    For iNumQtrs = 4 To 2 Step -1
        sMsg = "Enter sales data for" & Str(iNumQtrs) & " quarter(s)?"
        Form1.Visible = False
' Form1 is made invisible because msgbox opening will activate form window itself.
        iRet = MsgBox(sMsg, vbYesNo Or vbQuestion Or vbMsgBoxSetForeground, "Quarterly Sales")
'vbmsgboxsetforeground will activate form window itself
        If iRet = vbYes Then
            Exit For
        End If
    Next iNumQtrs
    sMsg = "Displaying data for" & Str(iNumQtrs) & " quarter(s)."
    MsgBox sMsg, vbMsgBoxSetForeground, "Quarterly Sales"
' Starting at E1, fill headers for the number of columns selected.
    Set oResizeRange = oWS.Range("E1").Resize(1, iNumQtrs)
'Range.resize returns a resized range which is referenced to the object variable oResizeRange
'Resized range contains 1 row and iNumQtrs number of columns
    oResizeRange.Formula = "=""Q"" & column()-4 & "" "" & ""Sales"""
' column() returns the current column number. 4 is subtracted from it to obtain the reqiured number
' Change the Orientation and WrapText properties for the headers.
    oResizeRange.WrapText = True
' Fill the interior color of the headers.
    oResizeRange.Interior.Color = vbRed
    oResizeRange.Font.Color = vbYellow
' Fill the columns with a formula and apply a number format.
    Set oResizeRange = oWS.Range("E2", "E6").Resize(, iNumQtrs)
    oResizeRange.Formula = "=RAND()*100"
    oResizeRange.NumberFormat = "[$$-409]#,##0.00"
    oResizeRange.Borders.Weight = xlThin
' Add a Totals formula for the sales data and apply a border.
    Set oResizeRange = oWS.Range("E8").Resize(, iNumQtrs)
    oResizeRange.Formula = "=SUM(E2:E6)"
    With oResizeRange.Borders(xlEdgeBottom)
                                            .LineStyle = xlDouble
                                            .Weight = xlThick
    End With
' Add a Chart for the selected data
    Set oResizeRange = oWS.Range("E2:E6").Resize(, iNumQtrs)
    Set oChart = oWS.Parent.Charts.Add
'oWS.Parent allows you to access oWS parent's property. ie. the workbook here
' here we add an entire new sheet which is a chart sheet
      With oChart
         .ChartWizard oResizeRange, xl3DColumn, , xlColumns
         .SeriesCollection(1).XValues = oWS.Range("A2", "A6")
            For iRet = 1 To iNumQtrs
               .SeriesCollection(iRet).Name = "=""Q" & Str(iRet) & """"
            Next iRet
         .Location xlLocationAsObject, oWS.Name
      End With
' Move the chart so as not to cover your data.
    With oWS.Shapes("Chart 1")
                            .Top = oWS.Rows(10).Top
                            .Left = oWS.Columns(2).Left
    End With
End Sub









