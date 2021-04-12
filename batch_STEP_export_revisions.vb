' Source: http://blog.ads-sol.com/2016/01/batch-assembly-export.html
' Takes assembly file and exports all parts as individual STEP files
' Inculdes part revisions
' Add new rule in iLogic and run within assy file
' Working inventor 2020

Dim oAsmDoc As AssemblyDocument
oAsmDoc = ThisApplication.ActiveDocument
oAsmName = ThisDoc.FileName(False) 'without extension

If ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
    MessageBox.Show("Please run this rule from the assembly file.", "iLogic")
    Exit Sub
End If
'get user input
RUsure = MessageBox.Show ( _
"This will create a STEP file for all components." _
& vbLf & " " _
& vbLf & "Are you sure you want to create STEP Drawings for all of the assembly components?" _
& vbLf & "This could take a while.", "iLogic - Batch Output STEPs ",MessageBoxButtons.YesNo)
If RUsure = vbNo Then
    Return
Else
End If
'- - - - - - - - - - - - -STEP setup - - - - - - - - - - - -
oPath = ThisDoc.Path
'get STEP target folder path
oFolder = oPath & "\" & oAsmName & " STEP Files"

'get the document revision to use in the new filename
oRevNumAsm = iProperties.Value("Project", "Revision Number")

'Check for the step folder and create it if it does not exist
If Not System.IO.Directory.Exists(oFolder) Then
System.IO.Directory.CreateDirectory(oFolder)
End If


'- - - - - - - - - - - - -Assembly - - - - - - - - - - - -
ThisDoc.Document.SaveAs(oFolder & "\" & oAsmName & "_Rev" & oRevNumAsm &(".stp") , True)

'- - - - - - - - - - - - -Components - - - - - - - - - - - -
'look at the files referenced by the assembly
Dim oRefDocs As DocumentsEnumerator
oRefDocs = oAsmDoc.AllReferencedDocuments
Dim oRefDoc As Document
'work the referenced models
For Each oRefDoc In oRefDocs
    Dim oCurFile As Document
    oCurFile = ThisApplication.Documents.Open(oRefDoc.FullFileName, True)
    oCurFileName = oCurFile.FullFileName

   
    'defines backslash As the subdirectory separator
    Dim strCharSep As String = System.IO.Path.DirectorySeparatorChar
   
    'find the postion of the last backslash in the path
    FNamePos = InStrRev(oCurFileName, "\", -1)  
    'get the file name with the file extension
    Name = Right(oCurFileName, Len(oCurFileName) - FNamePos)
    'get the file name (without extension)
    ShortName = Left(Name, Len(Name) - 4)
    oRevDoc = iProperties.Value(Name, "Project", "Revision Number")


    Try
        oCurFile.SaveAs(oFolder & "\" & ShortName & "_Rev" & oRevDoc & (".stp") , True)
    Catch
        MessageBox.Show("Error processing " & oCurFileName, "ilogic")
    End Try
    oCurFile.Close
Next
'- - - - - - - - - - - - -
MessageBox.Show("New Files Created in: " & vbLf & oFolder, "iLogic")
'open the folder where the new files are saved
Shell("explorer.exe " & oFolder,vbNormalFocus)