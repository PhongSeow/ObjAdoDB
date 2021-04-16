'**********************************
'* Name: pError
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.pError
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 18/2/2021
'* 1.0.2  16/4/2021	Remove excess Me.ClearErr()
'**********************************
Public Class pError
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.2"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public ReadOnly Property Description() As String
		Get
			Try
				Return Me.Obj.Description
			Catch ex As Exception
				Me.SetSubErrInf("Description.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property HelpContext() As Long
		Get
			Try
				Return Me.Obj.HelpContext
			Catch ex As Exception
				Me.SetSubErrInf("HelpContext.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property HelpFile() As String
		Get
			Try
				Return Me.Obj.HelpFile
			Catch ex As Exception
				Me.SetSubErrInf("HelpFile.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property NativeError() As Long
		Get
			Try
				Return Me.Obj.NativeError
			Catch ex As Exception
				Me.SetSubErrInf("NativeError.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Number() As Long
		Get
			Try
				Return Me.Obj.Number
			Catch ex As Exception
				Me.SetSubErrInf("Number.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Source() As String
		Get
			Try
				Return Me.Obj.Source
			Catch ex As Exception
				Me.SetSubErrInf("Source.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property SQLState() As String
		Get
			Try
				Return Me.Obj.SQLState
			Catch ex As Exception
				Me.SetSubErrInf("SQLState.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

End Class
