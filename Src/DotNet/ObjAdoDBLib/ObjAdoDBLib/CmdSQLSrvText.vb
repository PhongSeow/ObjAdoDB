'**********************************
'* Name: CmdSQLSrvText
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Command for SQL Server SQL statement Text
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.6
'* Create Time: 15/5/2021
'* 1.0.2	18/4/2021	Modify Execute,ParaValue
'* 1.0.3	17/5/2021	Modify ParaValue,ActiveConnection,Execute
'* 1.0.4	14/7/2021	Add DebugStr,mSQLStr
'* 1.0.5	15/7/2021	Add mSQLStr,mGetStr,ParaNameList Modify DebugStr
'* 1.0.6	18/7/2021	Modify DebugStr
'**********************************
Public Class CmdSQLSrvText
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.6"
	Public Property SQLText As String
	Private moCommand As Command

	Public Sub New(SQLText As String)
		MyBase.New(CLS_VERSION)
		Dim strStepName As String = ""
		Try
			Me.SQLText = SQLText
			moCommand = New Command
			With moCommand
				.CommandType = Command.CommandTypeEnum.adCmdText
				.CommandText = SQLText
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", strStepName, ex)
		End Try
	End Sub

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
				moCommand.Obj.ActiveConnection = value.Obj
				If moCommand.LastErr <> "" Then Throw New Exception(moCommand.LastErr)
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Set", ex)
			End Try
		End Set
	End Property

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum)
		Me.mAddPara(ParaName, DataType)
	End Sub


	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, Size As Long)
		Me.mAddPara(ParaName, DataType, Size)
	End Sub

	Private Sub mAddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, Optional Size As Long = -1)
		Dim strStepName As String = ""
		Try
			Dim oParameter As Parameter
			Dim dyeAny As Field.DataTypeEnum
			dyeAny = GetSQLSrvAdoDataType(DataType)
			Dim pdeAny As Parameter.ParameterDirectionEnum
			pdeAny = Parameter.ParameterDirectionEnum.adParamInput
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

	Public Function Execute() As Recordset
		Try
			Execute = New Recordset
			Execute.Obj = moCommand.Obj.Execute(mlngRecordsAffected)
			If moCommand.LastErr <> "" Then Throw New Exception(moCommand.LastErr)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Execute", ex)
			Return Nothing
		End Try
	End Function

	''' <summary>
	''' Records Affected by the execution of the Stored Procedure
	''' </summary>
	Private mlngRecordsAffected As Long
	Public ReadOnly Property RecordsAffected() As Long
		Get
			Return mlngRecordsAffected
		End Get
	End Property

	Public Property ParaValue(ParaName As String) As Object
		Get
			Try
				ParaValue = moCommand.Obj.Parameters(ParaName).Value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ParaValue.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				moCommand.Obj.Parameters(ParaName).Value = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ParaValue.Set", ex)
			End Try
		End Set
	End Property

	''' <summary>
	''' Returns debugging information for executing SQL statements
	''' </summary>
	Public ReadOnly Property DebugStr() As String
		Get
			Dim strStepName As String = ""
			Try
				Dim strDebugStr As String = Me.SQLText & vbCrLf
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

	Private mstrParaNameList As String = ""
	Public ReadOnly Property ParaNameList() As String
		Get
			Return mstrParaNameList
		End Get
	End Property

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

End Class
