Attribute VB_Name = "DirectoryTreeModule"
'This module contains this program's main code.
Option Explicit

Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As String) As Long


'This procedure returns the directories contained by the current directory.
Private Function GetDirectories() As String()
On Error GoTo ErrorTrap
Dim Directories() As String
Dim Item As String

   Item = Dir$("*.*", vbArchive Or vbDirectory Or vbHidden Or vbReadOnly Or vbSystem)
   Do Until Item = vbNullString
      If (GetAttr(Item) And vbDirectory) = vbDirectory Then
         If Not (Item = "." Or Item = "..") Then
            ChDir Item
            ChDir ".."
   
            If SafeArrayGetDim(Directories()) = 0 Then
               ReDim Directories(0 To 0) As String
            Else
               ReDim Preserve Directories(LBound(Directories()) To UBound(Directories()) + 1) As String
            End If
            Directories(UBound(Directories())) = Item
         End If
      End If
NextDirectory:
      If Item = vbNullString Then Exit Do
      Item = Dir$(, vbArchive Or vbDirectory Or vbHidden Or vbReadOnly Or vbSystem)
   Loop
   
   GetDirectories = Directories()
   Exit Function
   
ErrorTrap:
   Resume NextDirectory
End Function

'This procedure writes the directory tree starting at the specified root to the specified file.
Private Sub GetDirectoryTree(Root As String, FileH As Long)
Dim Directories() As String
Dim DirectoryIndex() As Long

   ChDrive Left$(Root, InStr(Root, ":"))
   ChDir Root
   
   ReDim DirectoryIndex(0 To 0) As Long
   DirectoryIndex(UBound(DirectoryIndex())) = 0
   Do
      DoEvents
      Do
         Directories() = GetDirectories()
         If SafeArrayGetDim(Directories()) = 0 Then Exit Do
         ChDir Directories(DirectoryIndex(UBound(DirectoryIndex())))
         ReDim Preserve DirectoryIndex(LBound(Directories()) To UBound(DirectoryIndex()) + 1) As Long
      Loop
   
      Do
         Print #FileH, CurDir$()
         ChDir ".."
         Directories() = GetDirectories()
         If UBound(DirectoryIndex()) = LBound(DirectoryIndex()) Then Exit Sub
         ReDim Preserve DirectoryIndex(LBound(DirectoryIndex()) To UBound(DirectoryIndex()) - 1) As Long
         
         If DirectoryIndex(UBound(DirectoryIndex())) < UBound(Directories()) Then
            DirectoryIndex(UBound(DirectoryIndex())) = DirectoryIndex(UBound(DirectoryIndex())) + 1
            Exit Do
         End If
      Loop
   Loop
End Sub

'This procedure scans through a directory tree.
Private Sub Main()
On Error GoTo ErrorTrap
Dim FileH As Long
Dim OutputFile As String
Dim Root As String

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   Root = Left$(App.Path, InStr(App.Path, ":")) & "\"
   OutputFile = App.Path
   If Not Right$(OutputFile, 1) = "\" Then OutputFile = OutputFile & "\"
   OutputFile = OutputFile & "Tree.txt"
   
   Root = InputBox$("Start at:", , Root)
   If Root = vbNullString Then Exit Sub
   OutputFile = InputBox$("Output file: ", , OutputFile)
   If OutputFile = vbNullString Then Exit Sub
   
   FileH = FreeFile()
   Open OutputFile For Output Lock Read Write As FileH
      GetDirectoryTree Root, FileH
   Close FileH
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   MsgBox Err.Description, vbExclamation
   Resume EndRoutine
End Sub

