'**********************************
'* Name: CmdSQLSrvSp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Command for SQL Server StoredProcedure
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 17/4/2021
'* 1.0.2	18/4/2021	Modify ActiveConnection
'* 1.0.3	24/4/2021	Add mAdoDataType
'* 1.0.4	25/4/2021	Modify New
'* 1.0.5	28/4/2021	Add ActiveConnection,AddPara,ParaValue,Execute
'* 1.0.6	16/5/2021	SQLSrvDataTypeEnum move to ConnSQLSrv, Modify Execute,ParaValue,ActiveConnection
'* 1.0.7	14/7/2021	Add DebugStr,mSQLStr,mGetStr,ParaNameList Modify New
'* 1.0.8	18/7/2021	Modify DebugStr
'* 1.1		1/10/2021	Add KeyName,CacheQuery
'* 1.2		20/12/2021	Modify CacheQuery
'**********************************
Imports PigKeyCacheLib
Public Class CmdSQLSrvSp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.2.2"
	Private moCommand As Command

	Public Sub New(SpName As String)
		MyBase.New(CLS_VERSION)
		Dim strStepName As String = ""
		Try
			Me.SpName = SpName
			moCommand = New Command
			With moCommand
				.CommandType = Command.CommandTypeEnum.adCmdStoredProc
				.CommandText = SpName
				Dim oParameter As Parameter
				strStepName = "CreateParameter(RETURN_VALUE)"
				oParameter = .CreateParameter("RETURN_VALUE", Field.DataTypeEnum.adInteger, Parameter.ParameterDirectionEnum.adParamReturnValue, 4)
				strStepName = "Append(RETURN_VALUE)"
				.Parameters.Append(oParameter)
				oParameter = Nothing
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", strStepName, ex)
		End Try
	End Sub

	''' <summary>
	''' Stored Procedure Name
	''' </summary>
	Private mstrSpName As String
	Public Property SpName() As String
		Get
			Return mstrSpName
		End Get
		Set(ByVal value As String)
			mstrSpName = value
		End Set
	End Property

	Private mstrParaNameList As String = ""
	Public ReadOnly Property ParaNameList() As String
		Get
			Return mstrParaNameList
		End Get
	End Property

	''' <summary>
	''' Returns debugging information for executing SQL statements
	''' </summary>
	Public ReadOnly Property DebugStr() As String
		Get
			Dim strStepName As String = ""
			Try
				strStepName = "SpName"
				Dim strDebugStr As String = "EXEC " & Me.SpName & " "
				Dim bolIsBegin As Boolean = False
				If Not moCommand.Parameters Is Nothing Then
					Dim strParaNameList As String = Me.ParaNameList
					Do While True
						If strParaNameList.Length <= 0 Then Exit Do
						Dim strParaName As String = Me.mGetStr(strParaNameList, "<", ">")
						If strParaName = "" Then Exit Do
						strStepName = "Parameters(" & strParaName & ")"
						strStepName &= "GetParameter"
						Dim oParameter As Parameter = moCommand.Parameters.Item(strParaName)
						If oParameter.LastErr <> "" Then Throw New Exception(oParameter.LastErr)
						With oParameter
							If .Direction <> Parameter.ParameterDirectionEnum.adParamReturnValue Then
								strStepName &= "GetValue"
								If bolIsBegin = True Then
									strDebugStr &= " , "
								Else
									bolIsBegin = True
								End If
								Dim strValue As String = ""
								Dim oValue As Object = .Value
								If Not oValue Is Nothing Then
									strValue = oValue.ToString
								End If
								strDebugStr &= .Name & "=" & mSQLStr(strValue)
							End If
						End With
					Loop
				End If
				Return strDebugStr
			Catch ex As Exception
				Me.SetSubErrInf("DebugStr", strStepName, ex)
				Return ""
			End Try
		End Get
	End Property

	Private Function mSQLStr(SrcValue As String, Optional IsNotNull As Boolean = False) As String
		SrcValue = Replace(SrcValue, "'", "''")
		If UCase(SrcValue) = "NULL" And IsNotNull = False Then
			mSQLStr = "NULL"
		Else
			mSQLStr = "'" & SrcValue & "'"
		End If
	End Function


	''' <summary>
	''' Stored Procedure return value
	''' </summary>
	Private mstrReturnValue As String
	Public ReadOnly Property ReturnValue() As Integer
		Get
			Return mstrReturnValue
		End Get
	End Property

	''' <summary>
	''' Records Affected by the execution of the Stored Procedure
	''' </summary>
	Private mlngRecordsAffected As Long
	Public ReadOnly Property RecordsAffected() As Long
		Get
			Return mlngRecordsAffected
		End Get
	End Property


	Public Function Execute() As Recordset
		Try
			Execute = moCommand.Execute(mlngRecordsAffected)
			If moCommand.LastErr <> "" Then Throw New Exception(moCommand.LastErr)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Execute", ex)
			Return Nothing
		End Try
	End Function

	Public Property ParaValue(ParaName As String) As Object
		Get
			Try
				ParaValue = moCommand.Parameters.Obj(ParaName).Value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ParaValue.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				moCommand.Parameters.Obj(ParaName).Value = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ParaValue.Set", ex)
			End Try
		End Set
	End Property

	Public Property ActiveConnection() As Connection
		Get
			Try
				Return moCommand.ActiveConnection
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Connection)
			Try
				moCommand.ActiveConnection = value
				If moCommand.LastErr <> "" Then Throw New Exception(moCommand.LastErr)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Set", ex)
			End Try
		End Set
	End Property

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum)
		Me.mAddPara(ParaName, DataType)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, IsOutPut As Boolean)
		Me.mAddPara(ParaName, DataType,, IsOutPut)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, Size As Long)
		Me.mAddPara(ParaName, DataType, Size)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, Size As Long, IsOutPut As Boolean)
		Me.mAddPara(ParaName, DataType, Size)
	End Sub

	Private Sub mAddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, Optional Size As Long = -1, Optional IsOutPut As Boolean = False)
		Dim strStepName As String = ""
		Try
			Dim oParameter As Parameter
			Dim dyeAny As Field.DataTypeEnum
			dyeAny = GetSQLSrvAdoDataType(DataType)
			Dim pdeAny As Parameter.ParameterDirectionEnum
			If IsOutPut = True Then
				pdeAny = Parameter.ParameterDirectionEnum.adParamOutput
			Else
				pdeAny = Parameter.ParameterDirectionEnum.adParamInput
			End If
			strStepName = "Append(" & ParaName & ")"
			oParameter = moCommand.CreateParameter(ParaName, dyeAny, pdeAny, Size)
			strStepName = "Append(" & ParaName & ")"
			moCommand.Parameters.Append(oParameter)
			oParameter = Nothing
			mstrParaNameList &= "<" & ParaName & ">"
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAddPara", strStepName, ex)
		End Try
	End Sub
	Private Function mGetStr(ByRef SourceStr As String, strBegin As String, strEnd As String, Optional IsCut As Boolean = True) As String
		Dim lngBegin As Long
		Dim lngEnd As Long
		Dim lngBeginLen As Long
		Dim lngEndLen As Long
		Try
			lngBeginLen = Len(strBegin)
			lngBegin = InStr(SourceStr, strBegin, CompareMethod.Text)
			lngEndLen = Len(strEnd)
			If lngEndLen = 0 Then
				lngEnd = Len(SourceStr) + 1
			Else
				lngEnd = InStr(lngBegin + lngBeginLen + 1, SourceStr, strEnd, CompareMethod.Text)
				If lngBegin = 0 Then Err.Raise(-1, , "lngBegin=0")
			End If
			If lngEnd <= lngBegin Then Err.Raise(-1, , "lngEnd <= lngBegin")
			If lngBegin = 0 Then Err.Raise(-1, , "lngBegin=0[2]")

			mGetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
			If IsCut = True Then
				SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
			End If
			Me.ClearErr()
		Catch ex As Exception
			mGetStr = ""
			Me.SetSubErrInf("mGetStr", ex)
		End Try
	End Function

	''' <summary>
	''' 用于缓存的键值名称|The name of the key value used for caching
	''' </summary>
	''' <param name="HeadPartName">键值名称前缀部分|Prefix part of key name</param>
	''' <returns></returns>
	Public ReadOnly Property KeyName(Optional HeadPartName As String = "") As String
		Get
			Try
				Dim oPigMD5 As New PigToolsLiteLib.PigMD5(Me.DebugStr, PigToolsLiteLib.PigMD5.enmTextType.UTF8)
				KeyName = oPigMD5.PigMD5
				If HeadPartName <> "" Then KeyName = HeadPartName & "." & KeyName
				oPigMD5 = Nothing
			Catch ex As Exception
				Me.SetSubErrInf("KeyName", ex)
				Return ""
			End Try
		End Get
	End Property

	''' <summary>
	''' The cache query returns Recordset.AllRecordset2JSon. Note that for SQL statements with updated data, using the cache query may have unpredictable results.
	''' </summary>
	''' <returns></returns>
	Public Function CacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, Optional CacheTime As Integer = 60, Optional TextType As PigToolsLiteLib.PigText.enmTextType = PigToolsLiteLib.PigText.enmTextType.Unicode) As String
		Dim strStepName As String = ""
		Try
			With ConnSQLSrv
				If .PigKeyValueApp Is Nothing Then
					strStepName = "InitPigKeyValue"
					.InitPigKeyValue()
					If .LastErr <> "" Then Throw New Exception(.LastErr)
				End If
				Dim strKeyName As String = Me.KeyName
				strStepName = "GetPigKeyValue"
				Dim oPigKeyValue As PigKeyValue = .PigKeyValueApp.GetPigKeyValue(strKeyName)
				If .PigKeyValueApp.LastErr <> "" Then Throw New Exception(.PigKeyValueApp.LastErr)
				Dim bolIsExec As Boolean = False
				If oPigKeyValue Is Nothing Then
					bolIsExec = True
				ElseIf oPigKeyValue.IsExpired = True Then
					bolIsExec = True
				End If
				If bolIsExec = True Then
					Dim rsAny As Recordset
					strStepName = "Execute"
					rsAny = Me.Execute
					If Me.LastErr <> "" Then Throw New Exception(.LastErr)
					strStepName = "New PigKeyValue"
					oPigKeyValue = New PigKeyValue(strKeyName, Now.AddSeconds(CacheTime), rsAny.AllRecordset2JSon, TextType, PigKeyValue.EnmSaveType.SaveSpace)
					If oPigKeyValue.LastErr <> "" Then Throw New Exception(oPigKeyValue.LastErr)
					strStepName = "PigKeyValueApp.SavePigKeyValue"
					.PigKeyValueApp.SavePigKeyValue(oPigKeyValue)
					If .PigKeyValueApp.LastErr <> "" Then Throw New Exception(.PigKeyValueApp.LastErr)
				End If
				CacheQuery = oPigKeyValue.StrValue
				oPigKeyValue = Nothing
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CacheQuery", strStepName, ex)
			Return ""
		End Try
	End Function

End Class
