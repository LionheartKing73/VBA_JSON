Attribute VB_Name = "Module1"
Sub Macro1()
 
Dim http As Object, JSON As Object, i As Integer, k As Integer


 
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://jsonplaceholder.typicode.com/posts", False
http.Send
Set JSON = ParseJson(http.responseText)


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
