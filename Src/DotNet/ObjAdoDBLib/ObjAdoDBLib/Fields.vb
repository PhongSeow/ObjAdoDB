'**********************************
'* Name: Fields
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Fields
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.4
'* Create Time: 18/2/2021
'* 1.0.2  20/2/2021   Modify Item
'* 1.0.3  21/2/2021   Modify Item fix bug
'* 1.0.4  16/4/2021	Remove excess Me.ClearErr()
'**********************************
Public Class Fields
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.4"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public Sub Clear()
		Try
			Me.Obj.Clear()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Clear", ex)
		End Try
	End Sub
	Public ReadOnly Property Count() As Long
		Get
			Try
				Return Me.Obj.Count
			Catch ex As Exception
				Me.SetSubErrInf("Count.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(Index) As Field
		Get
			Try
				Dim oField As New Field
				oField.Obj = Me.Obj.Item(Index)
				Return oField
			Catch ex As Exception
				Me.SetSubErrInf("Item.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
End Class
