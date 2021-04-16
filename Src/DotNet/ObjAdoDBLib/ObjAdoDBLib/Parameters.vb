'**********************************
'* Name: Parameter
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Errors
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'* Create Time: 18/2/2021
'* 1.0.2	3/3/2021	Modify Append
'* 1.0.3  16/4/2021	Remove excess Me.ClearErr()
'**********************************
Public Class Parameters
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.3"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public Sub Append(ObjParameter As Parameter)
		Try
			Me.Obj.Append(ObjParameter.Obj)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Append", ex)
		End Try
	End Sub
	Public Property Count() As Long
		Get
			Try
				Return Me.Obj.Count
			Catch ex As Exception
				Me.SetSubErrInf("Count.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.Count = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Count.Set", ex)
			End Try
		End Set
	End Property
	Public Sub Delete(Index)
		Try
			Me.Obj.Delete(Index)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Delete", ex)
		End Try
	End Sub
	Public Property Item(Index) As Parameter
		Get
			Try
				Return Me.Obj.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item(Index).Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Parameter)
			Try
				Me.Obj.Item(Index) = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Item.Set", ex)
			End Try
		End Set
	End Property
	Public Sub Refresh()
		Try
			Me.Obj.Refresh()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Refresh", ex)
		End Try
	End Sub
End Class
