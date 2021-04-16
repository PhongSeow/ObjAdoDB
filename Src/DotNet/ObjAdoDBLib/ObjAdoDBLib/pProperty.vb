'**********************************
'* Name: Property
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Property
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.2
'* Create Time: 15/2/2021
'* 1.0.2  16/4/2021	Remove excess Me.ClearErr()
'**********************************
Public Class pProperty
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.2"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
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
	Public ReadOnly Property Name() As String
		Get
			Try
				Return Me.Obj.Name
			Catch ex As Exception
				Me.SetSubErrInf("Name.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Type() As Field.DataTypeEnum
		Get
			Try
				Return Me.Obj.Type
			Catch ex As Exception
				Me.SetSubErrInf("Type.Get", ex)
				Return Nothing
			End Try
		End Get
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
		Set(value As Object)
			Try
				Me.Obj.Value = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Value.Set", ex)
			End Try
		End Set
	End Property
End Class
