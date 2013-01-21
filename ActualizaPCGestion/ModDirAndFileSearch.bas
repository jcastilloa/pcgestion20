Attribute VB_Name = "ModDirAndFileSearch"

Sub FilesSearch(DrivePath As String, Ext As String)
On Error Resume Next
    Dim XDir() As String
    Dim TmpDir As String
    Dim FFound As String
    Dim DirCount As Long
    Dim X As Long
    'Initialises Variables
    DirCount = 0
    ReDim XDir(0) As String
    XDir(DirCount) = ""


    If Right(DrivePath, 1) <> "\" Then
        DrivePath = DrivePath & "\"
    End If
    'Enter here the code for showing the pat
    '     h being
    'search. Example: Form1.label2 = DrivePa
    '     th
    'Search for all directories and store in
    '     the
    'XDir() variable


    DoEvents
        TmpDir = dir(DrivePath, vbDirectory)


        Do While TmpDir <> ""


            If TmpDir <> "." And TmpDir <> ".." Then


                If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then
                    XDir(DirCount) = DrivePath & TmpDir & "\"
                    DirCount = DirCount + 1
                    ReDim Preserve XDir(DirCount) As String
                End If
            End If
            TmpDir = dir
        Loop
        'Searches for the files given by extensi
        '     on Ext
        FFound = dir(DrivePath & Ext)


        Do Until FFound = ""
            'Code in here for the actions of the fil
            '     es found.
            'Files found stored in the variable FFou
            '     nd.
            'Example: Form1.list1.AddItem DrivePath
            '     & FFound
            FrmActualiza.ListView1.AddItem DrivePath & FFound
            FFound = dir
        Loop
        'Recursive searches through all sub dire
        '     ctories


        For X = 0 To (UBound(XDir) - 1)
            FilesSearch XDir(X), Ext
        Next X
        If FFound = "" Then
        'FrmActualiza.LblStatus.Caption = "Search For " & Ext & " Finished"
        End If
    End Sub



