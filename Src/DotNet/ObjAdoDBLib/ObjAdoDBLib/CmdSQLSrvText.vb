'**********************************
'* Name: CmdSQLSrvSp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Command for SQL Server SQL statement Text
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.2
'* Create Time: 15/5/2021
'* 1.0.2	18/4/2021	Modify Execute,ParaValue
'**********************************
Public Class CmdSQLSrvText
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.2"
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
				moCommand.ActiveConnection = value
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
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAddPara", strStepName, ex)
		End Try
	End Sub

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

End Class
