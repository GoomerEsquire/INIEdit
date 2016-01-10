Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices

Public Class INIFile

	Protected Class Group

		Protected keyArray As Key() = {}
		Protected n As String

		Sub New(name As String)
			n = name
		End Sub

		Public Overrides Function ToString() As String
			Return n
		End Function

		Friend ReadOnly Property Name As String
			Get
				Return n
			End Get
		End Property

		Friend Function GetKey(name As String, Optional create As Boolean = False) As Key

			Dim lname As String = LCase(name)
			Dim index As Integer = -1

			For i As Integer = 0 To keyArray.Count - 1
				If keyArray(i) Is Nothing Then
					index = i
					Exit For
				ElseIf LCase(keyArray(i).Name) = lname Then
					Return keyArray(i)
				End If
			Next

			If create Then
				Dim keyObj As New Key(Me, name)
				If index = -1 Then
					Array.Resize(keyArray, keyArray.Count + 1)
					index = keyArray.Count - 1
				End If
				keyArray(index) = keyObj
				Return keyObj
			End If

			Return Nothing

		End Function

		Friend Function GetKeys() As Key()

			Dim count As Integer = 0

			For Each keyObj As Key In keyArray
				If keyObj Is Nothing Then Exit For
				count += 1
			Next

			Dim newArray(count - 1) As Key
			count = 0

			For Each keyObj As Key In keyArray
				If keyObj Is Nothing Then Exit For
				newArray(count) = keyObj
				count += 1
			Next

			Return newArray

		End Function

		Friend ReadOnly Property AllKeys As Key()
			Get
				Return keyArray
			End Get
		End Property

	End Class

	Protected Class Key

		Protected n As String
		Protected v As String = String.Empty
		Protected g As Group

		Sub New(group As Group, name As String)
			g = group
			n = name
		End Sub

		Public Overrides Function ToString() As String
			Return n + "=" + v
		End Function

		Friend ReadOnly Property Name As String
			Get
				Return n
			End Get
		End Property

		Friend ReadOnly Property Group As Group
			Get
				Return g
			End Get
		End Property

		Friend Property Value As String
			Get
				Return v
			End Get
			Set(value As String)
				v = value
			End Set
		End Property

		Friend Sub Remove()

			For i As Integer = 0 To g.AllKeys.Count - 1
				If g.AllKeys(i) Is Nothing Then Exit For
				If g.AllKeys(i) Is Me Then
					g.AllKeys(i) = Nothing
					g.AllKeys.Move(i, g.AllKeys.Count - 1)
					Exit For
				End If
			Next

		End Sub

	End Class

	Protected curPath As String
	Protected groupArray As Group() = {}
	Protected autoFlush As Boolean = False
	Protected enc As System.Text.Encoding = System.Text.Encoding.UTF8

	''' <summary>
	''' Creates a new INIFile object.
	''' </summary>
	Public Sub New()
	End Sub

	''' <summary>
	''' Creates a new INIFile object.
	''' </summary>
	''' <param name="path">The path to the INI-File.</param>
	Public Sub New(path As String)
		curPath = path
	End Sub

	''' <summary>
	''' Creates a new INIFile object.
	''' </summary>
	''' <param name="path">The path to the INI-File.</param>
	''' <param name="encoding">The encoding of the file read/written.</param>
	Public Sub New(path As String, encoding As System.Text.Encoding)
		curPath = path
		enc = encoding
	End Sub

	''' <summary>
	''' Creates a new INIFile object.
	''' </summary>
	''' <param name="encoding">The encoding of the data passed.</param>
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

		Dim newString As New StringBuilder

		For g As Integer = 0 To groupArray.Count - 1
			Dim keys As Key() = groupArray(g).GetKeys
			If keys.Count = 0 Then Continue For
			If g > 0 Then newString.AppendLine()
			newString.AppendLine("[" + groupArray(g).Name + "]")
			For k As Integer = 0 To keys.Count - 1
				newString.AppendLine(keys(k).ToString)
			Next
		Next

		Return newString.ToString

	End Function

	''' <summary>
	''' Converts the string from the specified file path in INI format to use with this library.
	''' </summary>
	Public Overloads Sub Read()

		If curPath.Length = 0 Then
			Throw New InvalidOperationException("Filepath was not specified!")
			Return
		End If

		Dim data As String = File.ReadAllText(curPath, enc)

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

	''' <param name="data">Convert this string instead.</param>
	Public Overloads Sub Read(data As String)

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
	Public Function GetGroups() As String()

		Dim count As Integer = 0

		For Each groupObj As Group In groupArray
			If groupObj.GetKeys.Count = 0 Then Continue For
			count += 1
		Next

		Dim newArray(count - 1) As String
		count = 0

		For Each groupObj As Group In groupArray
			If groupObj.GetKeys.Count = 0 Then Continue For
			newArray(count) = groupObj.Name
			count += 1
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
			Throw New ArgumentNullException("Group string cannot be null or empty!")
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
			Throw New ArgumentNullException("Group string cannot be null or empty!")
			Exit Function
		End If

		If Key = Nothing OrElse Key.Length = 0 Then
			Throw New ArgumentNullException("Key string cannot be null or empty!")
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
			Throw New ArgumentNullException("Group string cannot be null or empty!")
			Return
		End If

		If Key = Nothing OrElse Key.Length = 0 Then
			Throw New ArgumentNullException("Key string cannot be null or empty!")
			Return
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
			Throw New ArgumentNullException("Group string cannot be null or empty!")
			Return
		End If

		If Key = Nothing OrElse Key.Length = 0 Then
			Throw New ArgumentNullException("Key string cannot be null or empty!")
			Return
		End If

		Dim groupObj As Group = GetGroup(Group)
		If Not groupObj Is Nothing Then
			Dim keyObj As Key = groupObj.GetKey(Key)
			If Not keyObj Is Nothing Then
				keyObj.Remove()
				If AutoSave Then Save()
			End If
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
			Throw New InvalidOperationException("Filepath was not specified!")
			Return
		End If

		File.WriteAllText(curPath, toINI, enc)

	End Sub

	''' <summary>
	''' Deletes the associated INI-File from the hard drive.
	''' </summary>
	''' <remarks></remarks>
	Public Sub Delete()

		If curPath.Length = 0 Then
			Throw New InvalidOperationException("Filepath was not specified!")
			Return
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

<HideModuleName> Module Extensions

	<Extension> Public Function Move(Of T)(ByRef a As T(), source As Integer, destination As Integer) As Boolean

		Dim sourceItem As T = a(source)

		If source < destination Then
			Array.ConstrainedCopy(a, source + 1, a, source, destination - source)
			a(destination) = sourceItem
		ElseIf source > destination Then
			Array.ConstrainedCopy(a, destination, a, destination + 1, source - destination)
			a(destination) = sourceItem
		Else
			Return False
		End If

		Return True

	End Function

End Module
