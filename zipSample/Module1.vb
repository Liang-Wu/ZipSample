Imports Ionic.Zip
Imports Ionic.FileSelector
Imports System.IO
Imports System.Xml
Imports System.Xml.XmlElement

Module Module1

    Sub Main()
        Dim str As String = MoodleCourseBackup_ToggleAuthAndEnrolment("C://Sample-Copy.mbz", "nologin", "guest")
    End Sub
    Public Function MoodleCourseBackup_ToggleAuthAndEnrolment(s_MBZFilePathAndName As String, s_TargetAuthMethod As String, s_TargetEnrolmentType As String) As String
        'verify the path of the s_MBZFilePathAndName   
        If Not File.Exists(s_MBZFilePathAndName) Then
            Return "File not found"
        End If
        'else if it exists check the auth
        Dim s_TargetAuthMethod_List As List(Of String) = New List(Of String)(New String() {"db", "nologin", "manual"})
        If Not s_TargetAuthMethod_List.Contains(s_TargetAuthMethod) Then
            Return "Invalid s_TargetAuthMethod"
        End If
        'validate the  s_TargetEnrolmentType
        Dim s_TargetEnrolmentType_List As List(Of String) = New List(Of String)(New String() {"Blank XML Key", "student", "teacher", "editingteacher", "pchair", "guest", "user"})
        If Not s_TargetEnrolmentType_List.Contains(s_TargetEnrolmentType) Then
            Return "Invalid s_TargetEnrolmentType"
        End If

        'else process the logical
        Dim user_fileName As String = "users.xml"
        Dim roles_fileName As String = "course/roles.xml"
        Dim tempDirectory_updateXMLBackup = "updateXMLBackup"
        Dim tempPathStr As String
        'fileter result containner -by filenames as filters
        Dim selection As Object = New Object()

        Try
            'get the two zip files by their names
            Using zip As ZipFile = ZipFile.Read(s_MBZFilePathAndName)
                selection = (From e In zip.Entries
                Where e.FileName = user_fileName Or e.FileName = roles_fileName
                                            Select e)

                'set up the extracting folder
                Dim pathStr As String = Path.GetDirectoryName(s_MBZFilePathAndName)
                tempPathStr = Path.Combine(New String() {pathStr, tempDirectory_updateXMLBackup})
                If Not Directory.Exists(tempPathStr) Then
                    Directory.CreateDirectory(tempPathStr)
                End If
                Dim selectionFileNameList As List(Of String) = New List(Of String)()
                ' loop the items and extract the item to the directory 
                For Each item As ZipEntry In selection
                    item.Extract(tempPathStr, ExtractExistingFileAction.OverwriteSilently)
                    selectionFileNameList.Add(item.FileName)
                Next
                'get the list of the path of new extracted xml files
                Dim userxml_path As String = tempPathStr + "\" + user_fileName
                Dim rolesxml_path As String = tempPathStr + "\" + "course\roles.xml"
                Dim pathList As List(Of String) = New List(Of String)(New String() {userxml_path, rolesxml_path})
                'get the result of manipulate xml files 
                Dim updateResultStr As String = MoodleCourseBackup_LoadAndModifyXMLs(pathList, s_TargetAuthMethod, s_TargetEnrolmentType)
                If Not updateResultStr = "success" Then
                    Return s_MBZFilePathAndName + ": extract file to backup successfully, but failed to update xmls"
                End If
                'remove the old file 1st and then add new modified files into zip
                For Each fileNameInZip As String In selectionFileNameList
                    zip.RemoveEntry(fileNameInZip)
                Next
                zip.AddFile(userxml_path, "")
                zip.AddFile(rolesxml_path, "course")
                zip.Save()
            End Using
            'remove the temp folder for the extracted xml files
            If Directory.Exists(tempPathStr) Then
                Directory.Delete(tempPathStr, True)
            End If
        Catch ex As Exception
            Return ex.Message
        End Try

        Return "success"
    End Function
    Public Function MoodleCourseBackup_LoadAndModifyXMLs(pathStr As List(Of String), s_TargetAuthMethod As String, s_TargetEnrolmentType As String) As String
        If pathStr.Count = 0 Then
            Return "no files"
        End If

        'load the xml files
        Dim doc_user As XmlDocument = New XmlDocument()
        Dim doc_roles As XmlDocument = New XmlDocument()
        Try
            For Each Str As String In pathStr
                If Str.Contains("users.xml") Then
                    'load xml document nodes
                    doc_user.Load(Str)
                    'get the xml nodes with tag auth
                    Dim userNodeList As XmlNodeList = doc_user.GetElementsByTagName("auth")
                    'manipulate the selected nodes with business logical
                    If (userNodeList.Count > 0) Then
                        For Each userNode As XmlNode In userNodeList
                            userNode.InnerText = s_TargetAuthMethod
                        Next
                    End If
                    'save file to the directory
                    doc_user.Save(Str)
                Else 'pathStr.Contains("roles.xml") Then
                    doc_roles.Load(Str)
                    'get the xml nodes with tag componnet
                    Dim rolesNodeList As XmlNodeList = doc_roles.GetElementsByTagName("component")
                    'manipulate the selected nodes with business logical
                    If (rolesNodeList.Count > 0) Then
                        For Each roleNode As XmlNode In rolesNodeList
                            Select Case s_TargetEnrolmentType
                                Case "Blank XML Key"
                                    roleNode.InnerText = String.Empty
                                Case "student"
                                    roleNode.InnerText = "enrol_student"
                                Case "teacher"
                                    roleNode.InnerText = "enrol_teacher"
                                Case "editingteacher"
                                    roleNode.InnerText = "enrol_editingteacher"
                                Case "pchair"
                                    roleNode.InnerText = "enrol_pchair"
                                Case "guest"
                                    roleNode.InnerText = "enrol_guest"
                                Case "user"
                                    roleNode.InnerText = "enrol_user"
                            End Select
                        Next
                    End If
                    'save file to the directory
                    doc_roles.Save(Str)
                End If
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
        Return "success"
    End Function
End Module
