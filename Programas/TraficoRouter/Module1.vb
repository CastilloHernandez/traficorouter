Module Module1

    
    Public Class EstadisticasMaquina : Implements IDisposable

        Dim sNombre As String
        Dim sMac As String
        Dim iPaquetesBajadaTotal As Integer
        Dim iPaquetesBajadaUltimo As Integer
        Dim iPaquetesSubidaTotal As Integer
        Dim iPaquetesSubidaUltimo As Integer
        Dim bTieneDatosBajada As Boolean
        Dim bTieneDatosSubida As Boolean

        Sub New()

        End Sub
        Public Property TieneDatosBajada() As Boolean
            Get
                Return bTieneDatosBajada
            End Get
            Set(ByVal value As Boolean)
                bTieneDatosBajada = value
            End Set
        End Property
        Public Property TieneDatosSubida() As Boolean
            Get
                Return bTieneDatosSubida
            End Get
            Set(ByVal value As Boolean)
                bTieneDatosSubida = value
            End Set
        End Property
        Public Property Nombre() As String
            Get
                Return sNombre
            End Get
            Set(ByVal value As String)
                sNombre = value
            End Set
        End Property
        Public Property Mac() As String
            Get
                Return sMac
            End Get
            Set(ByVal value As String)
                sMac = value
            End Set
        End Property
        Public Property PaquetesSubidaTotal() As Integer
            Get
                bTieneDatosSubida = False
                Return iPaquetesSubidaTotal
            End Get
            Set(ByVal value As Integer)
                If iPaquetesSubidaTotal > 0 Then
                    If value >= iPaquetesSubidaTotal Then
                        iPaquetesSubidaUltimo = value - iPaquetesSubidaTotal
                    Else
                        iPaquetesSubidaUltimo = value
                    End If
                End If
                bTieneDatosSubida = True
                iPaquetesSubidaTotal = value
            End Set
        End Property
        Public Property PaquetesSubidaUltimo() As Integer
            Get
                Return iPaquetesSubidaUltimo

            End Get
            Set(ByVal value As Integer)
                iPaquetesSubidaUltimo = value
            End Set
        End Property

        Public Property PaquetesBajadaTotal() As Integer
            Get
                bTieneDatosBajada = False
                Return iPaquetesBajadaTotal
            End Get
            Set(ByVal value As Integer)
                bTieneDatosBajada = True
                If iPaquetesBajadaTotal > 0 Then
                    If value >= iPaquetesBajadaTotal Then
                        iPaquetesBajadaUltimo = value - iPaquetesBajadaTotal
                    Else
                        iPaquetesBajadaUltimo = value
                    End If
                Else
                    bTieneDatosBajada = False
                End If

                iPaquetesBajadaTotal = value
            End Set
        End Property
        Public Property PaquetesBajadaUltimo() As Integer
            Get
                bTieneDatosBajada = False
                Return iPaquetesBajadaUltimo
            End Get
            Set(ByVal value As Integer)
                iPaquetesBajadaUltimo = value
            End Set
        End Property

        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: free unmanaged resources when explicitly called
                End If

                ' TODO: free shared unmanaged resources
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public cMaquinas As Collection
    Public iPuntosGrafica As Integer
    Public iIntervaloMuestreo As Integer
    Public iGrosorLinea As Integer
    Public iMuestras As Integer
    Sub CargarOpciones()
        iPuntosGrafica = Int(GetSetting("GraficaInternet", "Grafica", "PuntosGrafica", "500"))
        iIntervaloMuestreo = Int(GetSetting("GraficaInternet", "Grafica", "IntervaloMuestreo", "3000"))
        iGrosorLinea = Int(GetSetting("GraficaInternet", "Grafica", "GrosorLinea", "3"))
    End Sub
    Sub GuardarOpciones()
        SaveSetting("GraficaInternet", "Grafica", "PuntosGrafica", iPuntosGrafica)
        SaveSetting("GraficaInternet", "Grafica", "IntervaloMuestreo", iIntervaloMuestreo)
        SaveSetting("GraficaInternet", "Grafica", "GrosorLinea", igrosorlinea)
    End Sub
    Sub GuardarPosicion(ByVal Forma As Form)
        If Forma.WindowState = FormWindowState.Minimized Then
        Else
            SaveSetting("GraficaInternet", "Ventana", Forma.Name & "Opacidad", Forma.Opacity.ToString)
            SaveSetting("GraficaInternet", "Ventana", Forma.Name & "Estado", Forma.WindowState.ToString)
            SaveSetting("GraficaInternet", "Ventana", Forma.Name & "Top", Forma.Top)
            SaveSetting("GraficaInternet", "Ventana", Forma.Name & "Left", Forma.Left)
            SaveSetting("GraficaInternet", "Ventana", Forma.Name & "Width", Forma.Width)
            SaveSetting("GraficaInternet", "Ventana", Forma.Name & "Height", Forma.Height)
            SaveSetting("GraficaInternet", "Ventana", Forma.Name & "Topmost", -Forma.TopMost)
        End If
    End Sub
    Sub ColocarForm(ByVal Forma As Form)
        Dim iTop As Integer
        Dim iLeft As Integer
        Dim iWidth As Integer
        Dim iHeight As Integer
        Dim sOpacidad As Single
        Dim bSiempreEncima As Boolean

        sOpacidad = Val(GetSetting("GraficaInternet", "Ventana", Forma.Name & "Opacidad", "1"))
        iTop = Val(GetSetting("GraficaInternet", "Ventana", Forma.Name & "Top", "0"))
        iLeft = Val(GetSetting("GraficaInternet", "Ventana", Forma.Name & "Left", "0"))
        iWidth = Val(GetSetting("GraficaInternet", "Ventana", Forma.Name & "Width", "400"))
        iHeight = Val(GetSetting("GraficaInternet", "Ventana", Forma.Name & "Height", "300"))
        bSiempreEncima = -Val(GetSetting("GraficaInternet", "Ventana", Forma.Name & "Topmost", "0"))

        If iLeft > System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width - iWidth Then iLeft = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width - iWidth
        If iTop > System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height - iHeight Then iTop = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height - iHeight
        If iLeft < 0 Then iLeft = 0
        If iTop < 0 Then iTop = 0
        Forma.Opacity = sOpacidad
        Forma.Top = iTop
        Forma.Left = iLeft
        Forma.Width = iWidth
        Forma.Height = iHeight
        Forma.TopMost = bSiempreEncima
        If GetSetting("GraficaInternet", "Ventana", Forma.Name & "Estado", "Normal") = "Normal" Then
        Else
            Forma.WindowState = FormWindowState.Maximized
        End If

    End Sub

End Module
