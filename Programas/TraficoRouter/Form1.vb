
Public Class Form1

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        GuardarPosicion(Me)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim s As String
        Dim p As Integer
        Dim a() As String
        Dim c() As String
        Dim i As Integer
        Dim l As String
        Dim bTotalPaginas As Boolean


        ColocarForm(Me)

        BackgroundWorker1.WorkerReportsProgress = False

        iMuestras = 0
        cMaquinas = New Collection



        Chart1.ChartAreas.Clear()
        Chart1.Series.Clear()

        Chart1.ChartAreas.Add("Grafica")

        CargarOpciones()




        p = 1
        bTotalPaginas = False
        Do


            s = Cortar(getWeb("http://192.168.0.1/userRpm/WlanMacFilterRpm.htm?Page=" & Format(p)), "var wlanFilterList = new Array(", "0,0 );")
            If Len(s) Then
                p = p + 1
            End If

            a = Split(s, vbLf)
            For Each l In a
                If Len(l) Then
                    c = Split(l, ",")
                    If cMaquinas.Contains(c(0)) Then
                        bTotalPaginas = True
                    Else
                        Using m As New EstadisticasMaquina
                            m.Mac = c(0)
                            m.Nombre = c(4)
                            cMaquinas.Add(m, m.Mac)
                            Chart1.Series.Add(m.Mac)

                            Chart1.Series(m.Mac).ChartType = DataVisualization.Charting.SeriesChartType.Line
                            Chart1.Series(m.Mac).LegendText = m.Nombre
                            Chart1.Series(m.Mac).LegendToolTip = m.Mac
                            Chart1.Series(m.Mac).BorderWidth = iGrosorLinea

                            For i = 1 To iPuntosGrafica
                                Chart1.Series(m.Mac).Points.AddXY("*", 0)
                                Chart1.Series(m.Mac).Points(Chart1.Series(m.Mac).Points.Count - 1).IsEmpty = True
                            Next
                        End Using
                    End If


                End If
            Next


        Loop Until bTotalPaginas

        Timer1.Interval = iIntervaloMuestreo
        Timer1.Enabled = True








    End Sub
    Function Cortar(ByVal Cadena As String, ByVal Desde As String, ByVal Hasta As String) As String
        Dim i As Integer
        Dim f As Integer
        Dim s As String

        s = Cadena
        If Len(s) Then
            i = s.IndexOf(Desde) + Len(Desde)
            f = s.Substring(i).IndexOf(Hasta)
            s = s.Substring(i, f)
        End If
        Return s
    End Function
    Public Function getWeb(ByRef sURL As String) As String

        Debug.Print("<GetWeb")
        Using myWebClient As New System.Net.WebClient()

            Try
                Dim myCredentialCache As New System.Net.CredentialCache()
                Dim myURI As New Uri(sURL)

                myWebClient.Encoding = System.Text.Encoding.UTF8
                myWebClient.Credentials = New System.Net.NetworkCredential("admin", "admin")
                
                Debug.Print("GetWeb " & sURL & ">")
                Return myWebClient.DownloadString(myURI)
            Catch ex As Exception
                Debug.Print("GetWeb (Vacio) " & sURL & ">")
                Return ""
            End Try
        End Using

    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim sAhora As String

        If BackgroundWorker1.IsBusy Then
            sAhora = Format(Now(), "HH:mm")
            For Each m As EstadisticasMaquina In cMaquinas
                Chart1.Series(m.Mac).Points.AddXY(sAhora, 0)
                Chart1.Series(m.Mac).Points(Chart1.Series(m.Mac).Points.Count - 1).IsEmpty = True
            Next
        Else
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim s As String
        Dim i As Integer
        s = ""
        For i = 1 To 2
            s = s & Cortar(getWeb("http://192.168.0.1/userRpm/WlanStationRpm.htm?Page=" & Format(i)), "var hostList = new Array(", "0,0 );")
            System.Threading.Thread.Sleep(500)
        Next i
        e.Result = s
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim l As String
        Dim c() As String
        Dim sAhora As String
        Dim iTemp As Integer
        Dim sMaximo As String
        Dim iMaximo As String
        Static iContador As Integer

        Debug.Print("<RunWorkerCompleted")
        If Len(e.Result) Then
            For Each l In Split(e.Result, vbLf)
                Debug.Print(l)
                If Len(l) Then
                    c = Split(l, ",")
                    cMaquinas.Item(c(0)).paquetessubidatotal = Int(c(3))
                    cMaquinas.Item(c(0)).paquetesbajadatotal = Int(c(2))

                End If
            Next l

        End If

        sAhora = Format(Now(), "HH:mm")
        itemp = 0
        iMaximo = 0
        sMaximo = ""
        For Each m As EstadisticasMaquina In cMaquinas
            If m.TieneDatosBajada Then
                iMuestras = iMuestras + 1
                iTemp = m.PaquetesBajadaUltimo
                If iTemp > iMaximo Then
                    iMaximo = iTemp
                    sMaximo = m.Mac
                End If
                Chart1.Series(m.Mac).Points.AddXY(sAhora, iTemp)
                Chart1.Series(m.Mac).Points(Chart1.Series(m.Mac).Points.Count - 1).ToolTip = m.Nombre & vbCrLf & sAhora
                Chart1.Series(m.Mac).IsVisibleInLegend = True
            Else
                If iMuestras > 0 Then
                    Chart1.Series(m.Mac).IsVisibleInLegend = False     
                End If
                Chart1.Series(m.Mac).Points.AddXY(sAhora, 0)
                Chart1.Series(m.Mac).Points(Chart1.Series(m.Mac).Points.Count - 1).IsEmpty = True
            End If
        Next
        iContador = iContador + 1
        If iContador >= iPuntosGrafica Then
            iContador = 0
            If Len(sMaximo) Then
                Chart1.Series(sMaximo).Points(Chart1.Series(sMaximo).Points.Count - 1).Label = cMaquinas(sMaximo).Nombre
                Chart1.Series(sMaximo).Points(Chart1.Series(sMaximo).Points.Count - 1).LabelAngle = 30
            End If
        End If

        LimitarPuntosGrafica()
        Debug.Print("RunWorkerCompleted>")
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        For Each m As EstadisticasMaquina In cMaquinas

            Chart1.Series(m.Mac).Points.Clear()


        Next

    End Sub
    Private Sub LimitarPuntosGrafica()
     
        Debug.Print("<LimitarPuntosGrafica")

        For Each m As EstadisticasMaquina In cMaquinas
            If Chart1.Series(m.Mac).Points.Count > iPuntosGrafica Then
                Do
                    Chart1.Series(m.Mac).Points.RemoveAt(0)
                Loop Until Chart1.Series(m.Mac).Points.Count <= iPuntosGrafica

            End If

        Next
   

        Chart1.ChartAreas("Grafica").RecalculateAxesScale()
        Debug.Print("LimitarPuntosGrafica>")
    End Sub

    Private Sub TamañoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TamañoToolStripMenuItem.Click
        Dim i As Integer
        Do
            i = Val(InputBox("Tamaño de la grafica (en puntos graficados)", , Format(iPuntosGrafica)))
        Loop Until i > 0
        iPuntosGrafica = i
        GuardarOpciones()
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Me.Opacity = 0.95

    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        Me.Opacity = 0.9
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        Me.Opacity = 1
    End Sub

    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click
        Me.Opacity = 0.8

    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        Me.Opacity = 0.7

    End Sub

    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click
        Me.Opacity = 0.6

    End Sub

    Private Sub SiempreEncimaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SiempreEncimaToolStripMenuItem.Click
        Me.TopMost = Not Me.TopMost
    End Sub

    Private Sub IntervaloToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IntervaloToolStripMenuItem.Click
        Dim i As Integer
        Timer1.Enabled = False
        Do
            i = Val(InputBox("Intervalo de muestreo (1 a 60 segundos)", , Format(iIntervaloMuestreo / 1000)))
        Loop Until i > 0 And i <= 60
        iIntervaloMuestreo = i * 1000
        GuardarOpciones()
        Timer1.Interval = iIntervaloMuestreo
        Timer1.Enabled = True
    End Sub

    Private Sub GrosorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GrosorToolStripMenuItem.Click
        Dim i As Integer
        Do
            i = Val(InputBox("Grosor de la linea", , Format(iGrosorLinea)))
        Loop Until i > 0
        iGrosorLinea = i

        GuardarOpciones()

        For Each m As EstadisticasMaquina In cMaquinas
            Chart1.Series(m.Mac).BorderWidth = iGrosorLinea
            
        Next
    End Sub

    Private Sub Chart1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Chart1.Click
        Debug.Print(sender.ToString & vbTab & e.ToString)
    End Sub
End Class




