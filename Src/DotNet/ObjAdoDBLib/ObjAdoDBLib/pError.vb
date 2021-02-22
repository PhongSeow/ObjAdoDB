Public Class pError
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public ReadOnly Property Description() As String
		Get
			Try
				Return Me.Obj.Description
				Me.ClearErr()
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
				Me.ClearErr()
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
				Me.ClearErr()
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
				Me.ClearErr()
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
				Me.ClearErr()
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
				Me.ClearErr()
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
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("SQLState.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

End Class
