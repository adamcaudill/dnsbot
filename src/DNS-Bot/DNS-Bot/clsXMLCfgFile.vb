'/*
' * DNS-Bot
' *          Copyright 2004 - DNS-Bot Team
' *          See Copyright.txt & License.txt for details
' *
' *
' * This program is free software; you can redistribute it and/or modify
' * it under the terms of the GNU General Public License as published by
' * the Free Software Foundation; either version 1, or (at your option)
' * any later version.
' *
' * This program is distributed in the hope that it will be useful,
' * but WITHOUT ANY WARRANTY; without even the implied warranty of
' * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' * GNU General Public License for more details.
' *
' * You should have received a copy of the GNU General Public License
' * along with this program; if not, write to the Free Software
' * Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
' *
' */

Imports System.Xml

Public Class clsXMLCfgFile

    Dim Doc As New XmlDocument
    Dim FileName As String
    Dim doesExist As Boolean

    Public Sub New(ByVal aFileName As String)
        FileName = aFileName
        Try
            Doc.Load(aFileName)
            doesExist = True
        Catch ex As Exception
            If Err.Number = 53 Then
                Doc.LoadXml(("<configuration>" & "</configuration>"))
                Doc.Save(aFileName)
            End If
        End Try
    End Sub

    Public Function GetConfigInfo(ByVal aSection As String, ByVal aKey As String, ByVal aDefaultValue As String) As Collection
        ' <Added by: Adam at: 7/9/2004-03:30:32 on machine: BALLER-STA1>
        'Remove some not-so-friendly letters that makes XML choke
        aKey = aKey.Replace("!", "_").Replace("@", "_")
        ' </Added by: Adam at: 7/9/2004-03:30:32 on machine: BALLER-STA1>
        ' return immediately if the file didn't exist
        If doesExist = False Then
            Return New Collection
        End If
        If aSection = "" Then
            ' if aSection = "" then get all section names
            Return getchildren("")
        ElseIf aKey = "" Then
            ' if aKey = "" then get all keynames for the section
            Return getchildren(aSection)
        Else
            Dim col As New Collection
            col.Add(getKeyValue(aSection, aKey, aDefaultValue))
            Return col
        End If
    End Function

    Public Function WriteConfigInfo(ByVal aSection As String, ByVal aKey As String, ByVal aValue As String) As Boolean
        Dim node1 As XmlNode
        Dim node2 As XmlNode
        ' <Added by: Adam at: 7/9/2004-03:30:32 on machine: BALLER-STA1>
        'Remove some not-so-friendly letters that makes XML choke
        aKey = aKey.Replace("!", "_").Replace("@", "_")
        ' </Added by: Adam at: 7/9/2004-03:30:32 on machine: BALLER-STA1>
        If aKey = "" Then
            ' find the section, remove all its keys and remove the section
            node1 = (Doc.DocumentElement).SelectSingleNode("/configuration/" & aSection)
            ' if no such section, return True
            If node1 Is Nothing Then Return True
            ' remove all its children
            node1.RemoveAll()
            ' select its parent ("configuration")
            node2 = (Doc.DocumentElement).SelectSingleNode("configuration")
            ' remove the section
            node2.RemoveChild(node1)
        ElseIf aValue = "" Then
            ' find the section of this key
            node1 = (Doc.DocumentElement).SelectSingleNode("/configuration/" & aSection)
            ' return if the section doesn't exist
            If node1 Is Nothing Then Return True
            ' find the key
            node2 = (Doc.DocumentElement).SelectSingleNode("/configuration/" & aSection & "/" & aKey)
            ' return true if the key doesn't exist
            If node2 Is Nothing Then Return True
            ' remove the key
            If node1.RemoveChild(node2) Is Nothing Then Return False
        Else
            ' Both the Key and the Value are filled 
            ' Find the key
            node1 = (Doc.DocumentElement).SelectSingleNode("/configuration/" & aSection & "/" & aKey)
            If node1 Is Nothing Then
                ' The key doesn't exist: find the section
                node2 = (Doc.DocumentElement).SelectSingleNode("/configuration/" & aSection)
                If node2 Is Nothing Then
                    ' Create the section first
                    Dim e As Xml.XmlElement = Doc.CreateElement(aSection)
                    ' Add the new node at the end of the children of ("configuration")
                    node2 = Doc.DocumentElement.AppendChild(e)
                    ' return false if failure
                    If node2 Is Nothing Then Return False
                    ' now create key and value
                    e = Doc.CreateElement(aKey)
                    e.InnerText = aValue
                    ' Return False if failure
                    If (node2.AppendChild(e)) Is Nothing Then Return False
                Else
                    ' Create the key and put the value
                    Dim e As Xml.XmlElement = Doc.CreateElement(aKey)
                    e.InnerText = aValue
                    node2.AppendChild(e)
                End If
            Else
                ' Key exists: set its Value
                node1.InnerText = aValue
            End If
        End If
        ' Save the document
        Doc.Save(FileName)
    End Function

    Private Function getKeyValue(ByVal aSection As String, ByVal aKey As String, ByVal aDefaultValue As String) As String
        Dim node As XmlNode
        node = (Doc.DocumentElement).SelectSingleNode("/configuration/" & aSection & "/" & aKey)
        If node Is Nothing Then Return aDefaultValue
        Return node.InnerText
    End Function

    Private Function getchildren(ByVal aNodeName As String) As Collection
        Dim col As New Collection
        Dim node As XmlNode
        Try
            ' Select the root if the Node is empty
            If aNodeName = "" Then
                node = Doc.DocumentElement
            Else
                ' Select the node given
                node = Doc.DocumentElement.SelectSingleNode(aNodeName)
            End If
        Catch
        End Try
        ' exit with an empty collection if nothing here
        If node Is Nothing Then Return col
        ' exit with an empty colection if the node has no children
        If node.HasChildNodes = False Then Return col
        ' get the nodelist of all children
        Dim nodeList As XmlNodeList = node.ChildNodes
        Dim i As Integer
        ' transform the Nodelist into an ordinary collection
        For i = 0 To nodeList.Count - 1
            col.Add(nodeList.Item(i).Name)
        Next
        Return col
    End Function

End Class
