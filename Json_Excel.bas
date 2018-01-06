Attribute VB_Name = "Module11"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
Dim http As Object, JSON As Object, i As Integer, k As Integer


Dim FSO As New FileSystemObject
Dim JsonTS As TextStream
Set JsonTS = FSO.OpenTextFile("C:\work\11.14VBA\a.json", ForReading)
JsonText = JsonTS.ReadAll
JsonTS.Close
Set JSON = ParseJson(JsonText)


i = 2
For Each Item In JSON
    k = 0
    For Each Colum In Item
        k = k + 1
        Sheets(1).Cells(i, k).Value = Item(Colum)
        If i = 2 Then Sheets(1).Cells(1, k).Value = Colum
    Next

    i = i + 1
Next
MsgBox ("Complete!")

End Sub
