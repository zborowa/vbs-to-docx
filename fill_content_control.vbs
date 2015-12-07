' Accept arguments '
Set args = WScript.Arguments
' Tags/IDs used in the content control fields '
Dim wordIds
	wordIds = Array("idClient", "idLocation", "idDate", "idContact")
Dim oWord
Dim oDoc
Dim sDocPath

' If the next line gives an error continue with the line that comes next '
On Error Resume Next
	' Look if a Word object is already open '
	Set oApp = GetObject(, "Word.Application")

' If number of errors is higher or lower than 0 create a new Word object '
If Err.Number <> 0 Then
	Set oApp = CreateObject("Word.Application")
End If

' Specify the Word document and open it for the user (it does not have to be a .dotm file) '
sDocPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\template.dotm"
Set oDoc = oApp.Documents.Open(sDocPath)
oApp.Visible = True

' Iterate through all arguments and fill them in the ContentControls'
for I = 0 to args.Count -1
	oDoc.SelectContentControlsByTag(wordIds(I))(1).Range.Text = args(I)
Next
