Imports System.IO

Public Class INIFile

	Protected Class Group

		Protected keyArray As Key() = {}
		Protected n As String

		Sub New(name As String)
			n = name
		End Sub

		Public ReadOnly Property Name As String
			Get
				Return n
			End Get
		End Property

		Public Function GetKey(name As String, Optional create As Boolean = False) As Key

			Dim keyObj As Key = Nothing
			Dim lname As String = LCase(name)
			Dim empty As Integer = -1

			For i As Integer = 0 To keyArray.Count - 1
				If keyArray(i) Is Nothing Then
					If empty = -1 Then empty = i
					Continue For
				End If
				If LCase(keyArray(i).Name) = lname Then
					keyObj = keyArray(i)
					Exit For
				End If
			Next

			If create AndAlso keyObj Is Nothing Then
				keyObj = New Key(name)
				If empty > -1 Then
					keyArray(empty) = keyObj
				Else
					Array.Resize(keyArray, keyArray.Count + 1)
					keyArray(keyArray.Count - 1) = keyObj
				End If
			End If

			Return keyObj

		End Function

		Public Function GetKeys() As Key()

			Dim newArray As Key() = {}

			For Each keyObj As Key In keyArray
				If keyObj Is Nothing Then Continue For
				Array.Resize(newArray, newArray.Count + 1)
				newArray(newArray.Count - 1) = keyObj
			Next

			Return newArray

		End Function

		Public Sub RemoveKey(name As String)

			name = LCase(name)

			For i As Integer = 0 To keyArray.Count - 1
				If keyArray(i) Is Nothing Then Continue For
				If LCase(keyArray(i).Name) = name Then
					keyArray(i) = Nothing
					Exit For
				End If
			Next

		End Sub

	End Class

	Protected Class Key

		Protected n As String
		Protected v As String = String.Empty
		Sub New(name As String)
			n = name
		End Sub

		Public ReadOnly Property Name As String
			Get
				Return n
			End Get
		End Property

		Public Property Value As String
			Get
				Return v
			End Get
			Set(value As String)
				v = value
			End Set
		End Property

	End Class

	Protected curPath As String
	Protected groupArray As Group() = {}
	Protected autoFlush As Boolean = False
	Protected enc As System.Text.Encoding = System.Text.Encoding.UTF8

	''' <summary>
	''' Creates a new INIFile object.
	''' </summary>
	''' <remarks></remarks>
	Public Sub New()
	End Sub

	''' <summary>
	''' Creates a new INIFile object.
	''' </summary>
	''' <param name="path">The path to the INI-File.</param>
	''' <remarks></remarks>
	Public Sub New(path As String)
		curPath = path
	End Sub

	''' <summary>
	''' Creates a new INIFile object.
	''' </summary>
	''' <param name="path">The path to the INI-File.</param>
	''' <param name="encoding">The encoding of the file read/written.</param>
	''' <remarks></remarks>
	Public Sub New(path As String, encoding As System.Text.Encoding)
		curPath = path
		enc = encoding
	End Sub

	''' <summary>
	''' Creates a new INIFile object.
	''' </summary>
	''' <param name="encoding">The encoding of the data passed.</param>
	''' <remarks></remarks>
	Public Sub New(encoding As System.Text.Encoding)
		enc = encoding
	End Sub

	Private Function isGroup(Line As String) As Boolean

		If Line.Length > 1 AndAlso Line(0) = "[" And Line(Line.Length - 1) = "]" Then
			Return True
		End If
		Return False

	End Function

	Private Function getGroupName(Line As String) As String

		If Line.Length > 2 AndAlso isGroup(Line) Then
			Return Line.Substring(1, Line.Length - 2)
		End If

		Return Nothing

	End Function

	Private Function GetGroup(Name As String, Optional Create As Boolean = False) As Group

		Dim groupObj As Group = Nothing
		Dim lName As String = LCase(Name)

		For Each grp As Group In groupArray
			If LCase(grp.Name) = lName Then
				Return grp
			End If
		Next

		If Create AndAlso groupObj Is Nothing Then
			groupObj = New Group(Name)
			Array.Resize(groupArray, groupArray.Count + 1)
			groupArray(groupArray.Count - 1) = groupObj
		End If

		Return groupObj

	End Function

	Private Function toINI() As String

		Dim NewConfingString As String = ""

		For Each groupObj As Group In groupArray
			Dim keyArray As Key() = groupObj.GetKeys
			If keyArray.Count = 0 Then Continue For
			NewConfingString += "[" + groupObj.Name + "]" + vbCrLf
			For Each keyObj As Key In keyArray
				NewConfingString += keyObj.Name + "=" + keyObj.Value + vbCrLf
			Next
			NewConfingString += vbCrLf
		Next

		Return NewConfingString

	End Function

	''' <summary>
	''' Converts an INI-File to use with this library.
	''' </summary>
	''' <remarks></remarks>
	Public Sub ReadFile()

		Dim confFile As String = ""
		Dim curGroup As String = ""

		If File.Exists(curPath) Then
			confFile = File.ReadAllText(curPath, enc)
		End If

		For Each ConfLine As String In Split(confFile, vbCrLf)
			If ConfLine.Length = 0 Then Continue For
			If ConfLine.StartsWith("#") OrElse ConfLine.StartsWith(";") OrElse ConfLine.StartsWith("//") Then
				Continue For
			End If
			If isGroup(ConfLine) Then
				curGroup = getGroupName(ConfLine)
				Continue For
			End If
			Dim StrArray As String() = Split(ConfLine, "=")
			If StrArray.Count >= 2 Then
				GetGroup(curGroup, True).GetKey(StrArray(0), True).Value = String.Join("=", StrArray, 1, StrArray.Count - 1)
			End If
		Next

	End Sub

	''' <summary>
	''' Converts a string in INI format to use with this library.
	''' </summary>
	''' <param name="data">The data to convert.</param>
	''' <remarks></remarks>
	Public Sub SetData(data As String)

		Dim curGroup As String = ""

		For Each ConfLine As String In Split(data, vbCrLf)
			If ConfLine.Length = 0 Then Continue For
			If ConfLine.StartsWith("#") OrElse ConfLine.StartsWith(";") OrElse ConfLine.StartsWith("//") Then
				Continue For
			End If
			If isGroup(ConfLine) Then
				curGroup = getGroupName(ConfLine)
				Continue For
			End If
			Dim StrArray As String() = Split(ConfLine, "=")
			If StrArray.Count >= 2 Then
				GetGroup(curGroup, True).GetKey(StrArray(0), True).Value = String.Join("=", StrArray, 1, StrArray.Count - 1)
			End If
		Next

	End Sub

	''' <summary>
	''' Returns all groups.
	''' </summary>
	''' <returns>An array of string.</returns>
	''' <remarks></remarks>
	Public Function GetGroups() As String()

		Dim newArray As String() = {}

		For Each groupObj As Group In groupArray
			If groupObj.GetKeys.Count = 0 Then Continue For
			Array.Resize(newArray, newArray.Count + 1)
			newArray(newArray.Count - 1) = groupObj.Name
		Next

		Return newArray

	End Function

	''' <summary>
	''' Returns all keys in a group.
	''' </summary>
	''' <param name="Group">The group from which all keys are to be listed.</param>
	''' <returns>An array of string.</returns>
	''' <remarks>Will throw an exception if the group string is empty or null.</remarks>
	Public Function GetKeys(Group As String) As String()

		If Group Is Nothing OrElse Group.Length = 0 Then
			Throw New NullReferenceException("Group string cannot be null or empty!")
			Exit Function
		End If

		Dim newArray As String() = {}
		Dim groupObj As Group = GetGroup(Group)

		If Not groupObj Is Nothing Then
			Dim keys As Key() = groupObj.GetKeys
			Array.Resize(newArray, keys.Count)
			For i As Integer = 0 To keys.Count - 1
				newArray(i) = keys(i).Name
			Next
		End If

		Return newArray

	End Function

	''' <summary>
	''' Returns a value from a group and key.
	''' </summary>
	''' <param name="Group">The group from which the key and value is to be read.</param>
	''' <param name="Key">The key from which the value is to be read.</param>
	''' <param name="DefaultValue">Default value to return if the key or group do not exist.</param>
	''' <returns>A string value.</returns>
	''' <remarks>Will throw an exception if the group string is empty or null.</remarks>
	Public Function GetVal(Group As String, Key As String, Optional DefaultValue As String = Nothing) As String

		If Group Is Nothing OrElse Group.Length = 0 Then
			Throw New NullReferenceException("Group string cannot be null or empty!")
			Exit Function
		End If

		If Key = Nothing OrElse Key.Length = 0 Then
			Throw New NullReferenceException("Key string cannot be null or empty!")
			Exit Function
		End If

		Dim groupObj As Group = GetGroup(Group)
		If Not groupObj Is Nothing Then
			Dim keyObj As Key = groupObj.GetKey(Key)
			If Not keyObj Is Nothing Then
				Return keyObj.Value
			End If
		End If

		Return DefaultValue

	End Function

	''' <summary>
	''' Creates or overwrites a new value.
	''' </summary>
	''' <param name="Group">The group in which the key and value is to be saved.</param>
	''' <param name="Key">The key in which the value is to be saved.</param>
	''' <param name="NewValue">The value that is to be saved.</param>
	''' <remarks>Will throw an exception if the group string is empty or null.</remarks>
	Public Sub SetVal(Group As String, Key As String, NewValue As String)

		If Group = Nothing OrElse Group.Length = 0 Then
			Throw New NullReferenceException("Group string cannot be null or empty!")
			Exit Sub
		End If

		If Key = Nothing OrElse Key.Length = 0 Then
			Throw New NullReferenceException("Key string cannot be null or empty!")
			Exit Sub
		End If

		GetGroup(Group, True).GetKey(Key, True).Value = NewValue

		If AutoSave Then Save()

	End Sub

	''' <summary>
	''' Removes a key and its value.
	''' </summary>
	''' <param name="Group">The group from which the key is to be removed.</param>
	''' <param name="Key">The key that is to be removed.</param>
	''' <remarks>Will throw an exception if the group string is empty or null.</remarks>
	Public Sub RemoveVal(Group As String, Key As String)

		If Group = Nothing OrElse Group.Length = 0 Then
			Throw New NullReferenceException("Group string cannot be null or empty!")
			Exit Sub
		End If

		If Key = Nothing OrElse Key.Length = 0 Then
			Throw New NullReferenceException("Key string cannot be null or empty!")
			Exit Sub
		End If

		Dim groupObj As Group = GetGroup(Group)

		If Not groupObj Is Nothing Then
			groupObj.RemoveKey(Key)
			If AutoSave Then Save()
		End If

	End Sub

	''' <summary>
	''' Returns the path to the INIFile object.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public ReadOnly Property Path() As String

		Get
			Return curPath
		End Get

	End Property

	''' <summary>
	''' Writes the data of the INIFile object to the hard drive.
	''' </summary>
	''' <remarks></remarks>
	Public Sub Save()

		If curPath.Length = 0 Then
			Throw New NullReferenceException("Cannot save file without filepath!")
		End If

		File.WriteAllText(curPath, toINI, enc)

	End Sub

	''' <summary>
	''' Deletes the associated INI-File from the hard drive.
	''' </summary>
	''' <remarks></remarks>
	Public Sub Delete()

		If curPath.Length = 0 Then
			Throw New NullReferenceException("Cannot delete file without filepath!")
		End If

		File.Delete(curPath)

	End Sub

	''' <summary>
	''' Clears all data from the INI-File object.
	''' </summary>
	''' <remarks></remarks>
	Public Sub Clear()

		groupArray = {}

	End Sub

	''' <summary>
	''' Gets or sets whether the data will be written to the hard drive after every change. Default is False.
	''' </summary>
	''' <remarks>Turning this on will impact performance significantly.</remarks>
	Public Property AutoSave As Boolean

		Get
			Return autoFlush
		End Get
		Set(value As Boolean)
			autoFlush = value
		End Set

	End Property

	''' <summary>
	''' Gets the encoding of the INIFile object.
	''' </summary>
	''' <returns>A system.text.encoding object.</returns>
	Public ReadOnly Property FileEncoding As System.Text.Encoding
		Get
			Return enc
		End Get
	End Property

End Class