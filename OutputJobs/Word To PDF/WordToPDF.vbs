' WordToPDF.vbs - Word document to PDF conversion script for Alitum Designer OutputJobs
' Written and tested for Altium Designer 21.9.2 on Windows 10
'
' USAGE:
'   1) Add this script to your project
'   2) Open an OutputJob
'   3) Add new 'Script Output' 'Report Outputs'
'   4) Configure the output (right click, or ALT+ENTER)
'       a) Select the word document to convert
'       b) Click OK
'   5) Assign to file structure output container
'   6) Generate outputs, script takes ~15 seconds to run
'
' LIMITATIONS:
'   1. Can only convert 1 Word document
'   2. File structure outputs can only support 1 of these outputs, or will name all PDF files the same
'   3. File structures can support many of these outputs if and only if you use the original file name
'
' Copyright (C) 2021 Mitchell Overdick
'
' This program is free software: you can redistribute it and/or modify it under
' the terms of the GNU General Public License as published by the Free Software
' Foundation, either version 3 of the License, or (at your option) any later
' version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT ANY
' WARRANTY; without even the implied warranty of  MERCHANTABILITY or FITNESS FOR
' A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along with
' this program.  If not, see <http://www.gnu.org/licenses/>.
'
' Mitchell Overdick
' 12/16/2021

Function Configure(parameters)
    'MsgBox parameters,65,"Configure"
    Dim sProjectPath

    ' Parse parameters into dictionary
    Set paramsDict = ADParamParse(parameters)

    sProjectPath = ""
    ' If file parameter was set, use it
    if paramsDict.Exists("file") Then
        sProjectPath = paramsDict.Item("file")
    Else
        paramsDict.Add "file",""
    end if

    ' if sProjectPath is still blank, use project folder
    if sProjectPath = "" Then
        Set Workspace = GetWorkspace
        ' Setting to empty path freezes altium, use "WordDoc.docx"
        sProjectPath = GetParentPath(Workspace.DM_FocusedProject.DM_ProjectFullPath) + "\WordDoc.docx"
    end if

    set fso = CreateObject("Scripting.FileSystemObject")

    ' Open file dialog
    sFilter = "Word Document (*.docx)|*.docx|All Files (*)|*|"

    MyFile = GetFileDlgEx(sProjectPath,sFilter,"Select Document")

    ' Only update parameter if a file was selected
    if Myfile <> "" Then
        paramsDict.Item("file") = Myfile
    end if

    Configure = ADParamCompile(paramsDict)
End Function

' No clue how this works yet
Function PredictOutputFileNames(names)
    ' MsgBox names,65,"PredictOutputFileNames"
    Set paramsDict = ADParamParse(names)
    PredictOutputFileNames = GetBasename(paramsDict.Item("file")) + ".pdf"
End Function

' This routine is called when OutJob is run
Sub Generate(Parameters)
    'MsgBox Parameters,65,"Generate"
    Dim sFileName
    Dim sDestFile
    Set paramsDict = ADParamParse(parameters)

    ' Compile output file name
    if paramsDict.Item("TargetFileName") <> "" Then
        sFileName = paramsDict.Item("TargetFileName")
    Else
        sFileName = GetBasename(paramsDict.Item("file")) + ".pdf"
    End if
    sDestFile = paramsDict.Item("TargetFolder") + sFileName

    DocToPdf  paramsDict.Item("file"), sDestFile
End Sub

' Convert word doc to PDF
' https://stackoverflow.com/questions/8807153/vbscript-to-convert-word-doc-to-pdf
Function DocToPdf(docInputFile, pdfOutputFile)
  Dim fileSystemObject
  Dim wordApplication
  Dim wordDocument
  Dim wordDocuments
  Dim baseFolder

  Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
  Set wordApplication = CreateObject("Word.Application")
  Set wordDocuments = wordApplication.Documents

  docInputFile = fileSystemObject.GetAbsolutePathName(docInputFile)
  baseFolder = fileSystemObject.GetParentFolderName(docInputFile)

  If Len(pdfOutputFile) = 0 Then
    pdfOutputFile = fileSystemObject.GetBaseName(docInputFile) + ".pdf"
  End If

  If Len(fileSystemObject.GetParentFolderName(pdfOutputFile)) = 0 Then
    pdfOutputFile = baseFolder + "\" + pdfOutputFile
  End If

  ' Disable any potential macros of the word document.
  wordApplication.WordBasic.DisableAutoMacros

  'CanWriteFile(docInputFile)
  Set wordDocument = wordDocuments.Open(docInputFile, False, True)

  ' See http://msdn2.microsoft.com/en-us/library/bb221597.aspx
  ' See https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb256835(v=office.12)
  wordDocument.ExportAsFixedFormat pdfOutputFile, 17

  wordDocument.Close WdDoNotSaveChanges
  wordApplication.Quit WdDoNotSaveChanges

  Set wordApplication = Nothing
  Set fileSystemObject = Nothing

End Function

' https://stackoverflow.com/questions/38643487/vbscript-browseforfile-function-how-to-specify-file-types
Function GetFileDlgEx(sIniDir,sFilter,sTitle)
    sIniDir = Replace(sIniDir,"\","\\")
    Set oDlg = CreateObject("WScript.Shell").Exec("mshta.exe ""about:<object id=d classid=clsid:3050f4e1-98b5-11cf-bb82-00aa00bdce0b></object><script>moveTo(0,-9999);eval(new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).Read("&Len(sIniDir)+Len(sFilter)+Len(sTitle)+41&"));function window.onload(){var p=/[^\0]*/;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(p.exec(d.object.openfiledlg(iniDir,null,filter,title)));close();}</script><hta:application showintaskbar=no />""")
    oDlg.StdIn.Write "var iniDir='" & sIniDir & "';var filter='" & sFilter & "';var title='" & sTitle & "';"
    GetFileDlgEx = oDlg.StdOut.ReadAll
End Function

' Return root path of a file
Function GetParentPath(path)
    dim filesys, a
    Set filesys = CreateObject("Scripting.FileSystemObject")
    a = filesys.GetParentFolderName(path)
    GetParentPath = a
End Function

' Return file name with extension from a full path
Function GetFileName(path)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(path)
    GetFileName = objFSO.GetFileName(objFile)
End Function

' Conver file name.extension to just name
Function GetBasename(path)
    dim fso
    set fso = createobject("scripting.filesystemobject")
    GetBasename = fso.getbasename(path)
End Function

' Parse Altium params into dict
' expects parameters in this format: "param1=value1|param2=value2"
Function ADParamParse(parameters)
    Dim paramsSplit

    ' Initialize our empty Dictionary
    Set paramsDict = CreateObject("Scripting.Dictionary")
    paramsDict.CompareMode = vbTextCompare

    ' Smallest valid string is a=b
    if Len(parameters) > 3 Then
        ' Split argument into param=value array entries
        paramsSplit = Split(parameters,"|")

        ' Turn each parameter into a Dictionary key
        for each param in paramsSplit
            dim keyval
            ' Split can return empty params, so check that param is not length 0
            if Len(param) > 0 Then
                keyval = Split(param, "=")
                paramsDict.add keyval(0),keyval(1)
            end if
        next
    end if

    Set ADParamParse = paramsDict
End Function

' Compile Altium parameters into string
' Compiles parameters to this format: "param1=value1|param2=value2"
Function ADParamCompile(dictionary)
    Dim paramsString

    paramsString = ""
    for each key in dictionary
        paramsString = paramsString + key + "=" + dictionary.Item(key) + "|"
    next

    ADParamCompile = paramsString
End Function

' https://stackoverflow.com/questions/12300678/how-can-i-determine-if-a-file-is-locked-using-vbs
Function CanWriteFile(path)
    Dim oFso, oFile
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFso.OpenTextFile(path, 8, True)
    If Err.Number = 0 Then oFile.Close
end Function
