Attribute VB_Name = "FileUtil"
Option Explicit

Private fso As Object

' fso�I�u�W�F�N�g��������
Private Sub fso_init()
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

'�t�@�C���I���_�C�A���O����t�@�C����I��
Public Function getFileByDialog(Optional ByVal file_filter As String) As String
    Dim file_path As Variant
    
    If Not IsMissing(file_filter) Then
        file_path = Application.GetOpenFilename(FileFilter:=file_filter)
    Else
        file_path = Application.GetOpenFilename()
    End If
    
    If file_path = False Then
        getFileByDialog = vbNullString
    Else
        getFileByDialog = Application.GetOpenFilename()
    End If
End Function

'�t�H���_�I���_�C�A���O����t�H���_�̃p�X���擾
Public Function getFolderByDialog() As String
    Dim Shell As Object
    Dim mypath As Object
    
    '�t�H���_�I���_�C�A���O�\��
    Set Shell = CreateObject("Shell.Application")
    
    Set mypath = Shell.BrowseForFolder(&O0, "��������Ώۃt�H���_��I��ł�������", &H1 + &H10)
    
    If Not mypath Is Nothing Then
        getFolderByDialog = mypath.items.Item.Path
        Set mypath = Nothing
    Else
        getFolderByDialog = vbNullString
    End If
End Function

'����̃t�H���_�z���ɂ���t�@�C���i�t���p�X�j�̈ꗗ�𐶐�����
Public Function getFileListAsCollection(ByVal dir_name As String, ByVal filter As String, ByRef flist As Collection, Optional recursive As Boolean = True) As Collection
    Dim fname, subf As Variant
    Dim full_name  As String
    
    Call fso_init
    If flist Is Nothing Then Set flist = New Collection
        
    '�܂������̃f�B���N�g���̃t�@�C����ǉ�
    fname = Dir(dir_name & "\" & filter)
    
    Do While fname <> ""
        full_name = dir_name & "\" & fname
        If fso.FileExists(full_name) Then
            flist.Add full_name
        End If
        fname = Dir
    Loop
    
    If recursive = True Then
        '���̌�T�u�t�H���_�ɂ��čċA�I�Ɏ��s
        For Each subf In fso.GetFolder(dir_name).SubFolders
            Set flist = getFileListAsCollection(subf.Path, filter, flist, True)
        Next
    End If
    
    Set getFileListAsCollection = flist

End Function

'����������΃f�B���N�g�����쐬����
Public Sub createFolderIfNotExists(ByVal folder As Variant)
    Call fso_init
    If Not fso.FolderExists(fso.GetParentFolderName(folder)) Then
        createFolderIfNotExists (fso.GetParentFolderName(folder))
    End If
    If Not fso.FolderExists(folder) Then
        fso.CreateFolder (folder)
    End If
End Sub

Public Function getFilenameFromFullpath(ByVal full_path As String)
    Dim pos As Long
    pos = InStrRev(full_path, "\")
    getFilenameFromFullpath = Mid(full_path, pos + 1)
End Function

Public Function getPathFromFullpath(ByVal full_path As String)
    Dim pos As Long
    pos = InStrRev(full_path, "\")
    getPathFromFullpath = Left(full_path, pos)
End Function

Public Function splitPath(ByVal full_path As String) As String()
    splitPath = Split(full_path, "\")
End Function
