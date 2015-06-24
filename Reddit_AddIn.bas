Attribute VB_Name = "Reddit_AddIn"

Option Explicit

Sub Convert_Selection_To_Reddit_Table()
'iterators
Dim i As Integer
Dim j As Integer
Dim k As Integer

'strings used for formatting and output
Dim formatString As String
Dim revFormatStr As String
Dim tempString As String
Dim FinalString As String
Dim cleanString As String

'helper measures
Dim tableRows As Integer
Dim tableCols As Integer

Dim MatrixArray As Range: Set MatrixArray = Selection
tableRows = MatrixArray.Rows.Count
tableCols = MatrixArray.Columns.Count

'The backslash MUST be the first character, else it will double up all of the slashes
cleanString = "\^*~`"



 
'Compiler directives are used to test for OS
#If Mac Then
    Dim objClipboard As MSForms.DataObject: Set objClipboard = New MSForms.DataObject
#ElseIf win64 Or Win32 Then
    Dim objClipboard As Object: Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
#Else
    MsgBox "Unknown operating system!"
    Exit Sub
#End If
 
If tableRows < 2 Or tableCols < 2 Then
        MsgBox "Selection Too Small, must be at least 2x2"
        Exit Sub
End If

'For each row
For i = 1 To tableRows
        If i = 2 Then 'Set the alignment and table formatting for the table. Based on the alignment of the second row of the selected table
                For j = 1 To tableCols
                        Select Case MatrixArray(1, j).HorizontalAlignment
                                Case xlGeneral, 1: FinalString = FinalString & "-:|" ' General
                                Case xlLeft, -4131: FinalString = FinalString & ":-|" ' Left
                                Case xlCenter, -4108: FinalString = FinalString & ":-:|" ' Center
                                Case xlRight, -4152: FinalString = FinalString & "-:| " ' Right
                                Case Else
                        End Select
                Next
                FinalString = FinalString & Chr(10)
        End If
        'For each column
        For j = 1 To tableCols
            'Using .Text here so that the formatted Excel display is used, instead of the underlying .Value
            tempString = MatrixArray(i, j).Text
            For k = 1 To Len(cleanString) 'escape characters are escaped. add characters in variable definition above
                If InStr(tempString, Mid(cleanString, k, 1)) > 0 Then tempString = Replace(tempString, Mid(cleanString, k, 1), "\" & Mid(cleanString, k, 1))
            Next k
                'Reddit formatting
                If MatrixArray(i, j).Font.Strikethrough Then
                    formatString = formatString & "~~" 'StrikeThrough
                    revFormatStr = "~~" & revFormatStr
                End If
                If MatrixArray(i, j).Font.Bold Then
                    formatString = formatString & "**" 'Bold
                    revFormatStr = "**" & revFormatStr 'Bold
                End If
                If MatrixArray(i, j).Font.Italic Then
                    formatString = formatString & "*" 'Italic
                    revFormatStr = "*" & revFormatStr
                End If
                If MatrixArray(i, j).Font.Superscript Then
                    formatString = formatString & "^" 'SuperScript
                End If
                
                'Build the cell contents
                FinalString = FinalString & formatString & tempString & revFormatStr & "|"
                formatString = vbNullString 'Clear format
                revFormatStr = vbNullString
        Next
        FinalString = FinalString & Chr(10)
Next

        'Max chars in Reddit comments is 10k. Hope you only wanted the table!
        If Len(FinalString) > 10000 Then
            MsgBox ("There are too many characters for Reddit comment! 10 000 characters copied.")
            FinalString = Left(FinalString, 9999)
        End If


  objClipboard.SetText FinalString
  objClipboard.PutInClipboard
 
Set MatrixArray = Nothing
Set objClipboard = Nothing

MsgBox "Data copied to clipboard!", vbOKOnly, "Written by: /u/norsk & /u/BornOnFeb2nd & " & Chr(10) & "/u/paxperscientiam"

End Sub
