' ===============================================================================
' Microsoft Configuration Management Application Block for .NET
' http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp
'
' XmlSerializableHashtable.vb
'
' A helper class used to serialize a Hashtable instance on Xml.
'
' For more information see the Configuration Management Application Block Implementation Overview. 
' 
' ===============================================================================
' Copyright (C) 2000-2001 Microsoft Corporation
' All rights reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
' OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
' FITNESS FOR A PARTICULAR PURPOSE.
' ==============================================================================

Imports System
Imports System.Collections
Imports System.Xml
Imports System.Xml.Serialization

Imports [MABCM] = Microsoft.ApplicationBlocks.ConfigurationManagement

' <summary>
' This class is used to provide serialization support for a hashtable
' Thanks to Christoph, christophdotnet@austin.rr.com, for the xml-serializable hashtable implementation 
' </summary>
<XmlInclude(GetType(String)), _
 XmlInclude(GetType(Boolean)), _
 XmlInclude(GetType(Short)), _
 XmlInclude(GetType(Integer)), _
 XmlInclude(GetType(Long)), _
 XmlInclude(GetType(Single)), _
 XmlInclude(GetType(Double)), _
 XmlInclude(GetType(DateTime)), _
 XmlInclude(GetType(Char)), _
 XmlInclude(GetType(Decimal)), _
 XmlInclude(GetType(UInt16)), _
 XmlInclude(GetType(UInt32)), _
 XmlInclude(GetType(UInt64)), _
 XmlInclude(GetType(Int64))> _
Public Class XmlSerializableHashtable

#Region "Nested Class--Item"

    ' <summary>
    ' Represents an entry for the hashtable
    ' </summary>

    Public Class Entry
        Private _entryKey As Object
        Private _entryValue As Object


        ' <summary>
        ' Default constructor, needed by serialization support
        ' </summary>
        Public Sub New()
        End Sub 'New

        ' <summary>
        ' Construct the Entity specifying the key and the entry
        ' </summary>
        ' <param name="entryKey"></param>
        ' <param name="entryValue"></param>
        Public Sub New(ByVal entryKey As Object, ByVal entryValue As Object)
            _entryKey = entryKey
            _entryValue = entryValue
        End Sub 'New

        ' <summary>
        ' Return the key
        ' </summary>

        <XmlElement("key")> _
        Public Property EntryKey() As Object
            Get
                Return _entryKey
            End Get
            Set(ByVal Value As Object)
                _entryKey = Value
            End Set
        End Property
        ' <summary>
        ' Return the entry value
        ' </summary>

        <XmlElement("value")> _
        Public Property EntryValue() As Object
            Get
                Return _entryValue
            End Get
            Set(ByVal Value As Object)
                _entryValue = Value
            End Set
        End Property
    End Class 'Entry 
#End Region

#Region "Declarations"

    Private _ht As Hashtable

#End Region


#Region "Constructors"


    ' <summary>
    ' Default constructor
    ' </summary>
    Public Sub New()
        _ht = New Hashtable(10)
    End Sub 'New



    ' <summary>
    ' Creates a serializable hashtable Imports a hashtable
    ' </summary>
    ' <param name="ht"></param>
    Public Sub New(ByVal ht As Hashtable)
        _ht = ht
    End Sub 'New


#End Region


#Region "Public Methods & Properties"

    ' <summary>
    ' Returns the contained hashtable
    ' </summary>

    <XmlIgnore()> _
    Public ReadOnly Property InnerHashtable() As Hashtable
        Get
            Return _ht
        End Get
    End Property
    ' <summary>
    ' Used to serilalize the contents of the hashtable
    ' </summary>

    Public Property Entries() As Entry()
        Get
            Dim entryArray(_ht.Count - 1) As Entry
            Dim i As Integer = 0

            Dim de As DictionaryEntry
            For Each de In _ht
                entryArray(i) = New Entry(de.Key, de.Value)
                i = i + 1
            Next de

            Return entryArray
        End Get

        Set(ByVal Value As Entry())
            SyncLock _ht.SyncRoot
                _ht.Clear()
                Dim item As Entry
                For Each item In Value
                    _ht.Add(GetValueFromXml(item.EntryKey), GetValueFromXml(item.EntryValue))
                Next item
            End SyncLock
        End Set
    End Property

#End Region
#Region "Private Methods"

    Private Function GetValueFromXml(ByVal value As Object) As Object
        Dim valueList As Object() = CType(value, Object())
        If (valueList.Length = 1) AndAlso (TypeOf valueList(0) Is XmlCharacterData) Then
            Return (CType(valueList(0), XmlCharacterData)).Value
        ElseIf (valueList.Length = 2) AndAlso (TypeOf valueList(1) Is XmlCharacterData) Then
            Return (CType(valueList(1), XmlCharacterData)).Value
        Else
            Return ""
        End If
    End Function 'GetValueFromXml

#End Region
End Class 'XmlSerializableHashtable 
