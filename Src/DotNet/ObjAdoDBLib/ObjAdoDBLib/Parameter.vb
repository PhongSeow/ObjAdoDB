﻿'**********************************
'* Name: Parameter
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Errors
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.4
'* Create Time: 18/2/2021
'* 1.0.2  3/3/2021 Modify New
'* 1.0.3  16/4/2021	Remove excess Me.ClearErr(), Modify New
'* 1.0.4  17/4/2021	Add Value
'**********************************
Public Class Parameter
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.4"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
		Try
			Me.Obj = CreateObject("ADODB.Parameter")
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", ex)
		End Try
	End Sub
	Public Enum ParameterDirectionEnum
		adParamInput = 1
		adParamInputOutput = 3
		adParamOutput = 2
		adParamReturnValue = 4
		adParamUnknown = 0
	End Enum
	Public Property Attributes() As Long
		Get
			Try
				Return Me.Obj.Attributes
			Catch ex As Exception
				Me.SetSubErrInf("Attributes.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.Attributes = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Attributes.Set", ex)
			End Try
		End Set
	End Property
	Public Property Direction() As ParameterDirectionEnum
		Get
			Try
				Return Me.Obj.Direction
			Catch ex As Exception
				Me.SetSubErrInf("Direction.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As ParameterDirectionEnum)
			Try
				Me.Obj.Direction = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Direction.Set", ex)
			End Try
		End Set
	End Property
	Public Property Name() As String
		Get
			Try
				Return Me.Obj.Name
			Catch ex As Exception
				Me.SetSubErrInf("Name.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.Name = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Name.Set", ex)
			End Try
		End Set
	End Property
	Public Property NumericScale() As Byte
		Get
			Try
				Return Me.Obj.NumericScale
			Catch ex As Exception
				Me.SetSubErrInf("NumericScale.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Byte)
			Try
				Me.Obj.NumericScale = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("NumericScale.Set", ex)
			End Try
		End Set
	End Property
	Public Property Value() As Object
		Get
			Try
				Return Me.Obj.Value
			Catch ex As Exception
				Me.SetSubErrInf("Value.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(Value As Object)
			Try
				Me.Obj.Value = Value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Value.Set", ex)
			End Try
		End Set
	End Property
End Class
