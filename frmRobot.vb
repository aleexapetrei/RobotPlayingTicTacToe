Public Class frmRobot
    Inherits System.Windows.Forms.Form
    Private m_TicTacToeGame As New TicTacToe
    Private m_iCurrentPlayer As TicTacToe.GridEntry = TicTacToe.GridEntry.PlayerX
    Private m_TicTacToeAI As New TicTacToeAI(TicTacToe.GridEntry.PlayerO)
    Private m_bIsAIEnabled As Boolean = True
    Private m_bGameEnded As Boolean = False
    Dim posicionFicha As String

    Private m_iPlayerXCount As Integer
    Private m_iPlayerOCount As Integer
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents pctJuego As System.Windows.Forms.PictureBox
    Friend WithEvents TimMuestreo As System.Windows.Forms.Timer
    Friend WithEvents PicResultado As System.Windows.Forms.PictureBox
    Friend WithEvents WebCam1 As WebCAM.WebCam
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Serial As System.IO.Ports.SerialPort
    Private m_iDrawCount As Integer

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'Grid Clicks
        AddHandler lblGrid00.Click, AddressOf lblGrid_Click
        AddHandler lblGrid01.Click, AddressOf lblGrid_Click
        AddHandler lblGrid02.Click, AddressOf lblGrid_Click
        AddHandler lblGrid10.Click, AddressOf lblGrid_Click
        AddHandler lblGrid11.Click, AddressOf lblGrid_Click
        AddHandler lblGrid12.Click, AddressOf lblGrid_Click
        AddHandler lblGrid20.Click, AddressOf lblGrid_Click
        AddHandler lblGrid21.Click, AddressOf lblGrid_Click
        AddHandler lblGrid22.Click, AddressOf lblGrid_Click

        Me.DrawGrid()

        AddHandler m_TicTacToeGame.TicTacToeWinOccured, AddressOf WeHaveAWinner
        AddHandler m_TicTacToeGame.Cat, AddressOf Cat

        MenuItem2.Checked = m_bIsAIEnabled

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblGrid22 As System.Windows.Forms.Label
    Friend WithEvents lblGrid21 As System.Windows.Forms.Label
    Friend WithEvents lblGrid20 As System.Windows.Forms.Label
    Friend WithEvents lblGrid12 As System.Windows.Forms.Label
    Friend WithEvents lblGrid11 As System.Windows.Forms.Label
    Friend WithEvents lblGrid10 As System.Windows.Forms.Label
    Friend WithEvents lblGrid02 As System.Windows.Forms.Label
    Friend WithEvents lblGrid01 As System.Windows.Forms.Label
    Friend WithEvents lblGrid00 As System.Windows.Forms.Label
    Friend WithEvents lblCurrentPlayer As System.Windows.Forms.Label
    Friend WithEvents lblYw As System.Windows.Forms.Label
    Friend WithEvents lblYourWins As System.Windows.Forms.Label
    Friend WithEvents lblMw As System.Windows.Forms.Label
    Friend WithEvents lblMyWins As System.Windows.Forms.Label
    Friend WithEvents lblD As System.Windows.Forms.Label
    Friend WithEvents lblDraws As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblGrid22 = New System.Windows.Forms.Label
        Me.lblGrid21 = New System.Windows.Forms.Label
        Me.lblGrid20 = New System.Windows.Forms.Label
        Me.lblGrid12 = New System.Windows.Forms.Label
        Me.lblGrid11 = New System.Windows.Forms.Label
        Me.lblGrid10 = New System.Windows.Forms.Label
        Me.lblGrid02 = New System.Windows.Forms.Label
        Me.lblGrid01 = New System.Windows.Forms.Label
        Me.lblGrid00 = New System.Windows.Forms.Label
        Me.lblCurrentPlayer = New System.Windows.Forms.Label
        Me.lblYw = New System.Windows.Forms.Label
        Me.lblYourWins = New System.Windows.Forms.Label
        Me.lblMw = New System.Windows.Forms.Label
        Me.lblMyWins = New System.Windows.Forms.Label
        Me.lblD = New System.Windows.Forms.Label
        Me.lblDraws = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.Label10 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.pctJuego = New System.Windows.Forms.PictureBox
        Me.TimMuestreo = New System.Windows.Forms.Timer(Me.components)
        Me.PicResultado = New System.Windows.Forms.PictureBox
        Me.WebCam1 = New WebCAM.WebCam
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Serial = New System.IO.Ports.SerialPort(Me.components)
        Me.GroupBox1.SuspendLayout()
        CType(Me.pctJuego, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicResultado, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblGrid22)
        Me.GroupBox1.Controls.Add(Me.lblGrid21)
        Me.GroupBox1.Controls.Add(Me.lblGrid20)
        Me.GroupBox1.Controls.Add(Me.lblGrid12)
        Me.GroupBox1.Controls.Add(Me.lblGrid11)
        Me.GroupBox1.Controls.Add(Me.lblGrid10)
        Me.GroupBox1.Controls.Add(Me.lblGrid02)
        Me.GroupBox1.Controls.Add(Me.lblGrid01)
        Me.GroupBox1.Controls.Add(Me.lblGrid00)
        Me.GroupBox1.Location = New System.Drawing.Point(412, 57)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 144)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Grid"
        '
        'lblGrid22
        '
        Me.lblGrid22.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid22.Location = New System.Drawing.Point(128, 104)
        Me.lblGrid22.Name = "lblGrid22"
        Me.lblGrid22.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid22.TabIndex = 17
        Me.lblGrid22.Text = "X"
        Me.lblGrid22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid21
        '
        Me.lblGrid21.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid21.Location = New System.Drawing.Point(88, 104)
        Me.lblGrid21.Name = "lblGrid21"
        Me.lblGrid21.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid21.TabIndex = 16
        Me.lblGrid21.Text = "X"
        Me.lblGrid21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid20
        '
        Me.lblGrid20.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid20.Location = New System.Drawing.Point(48, 104)
        Me.lblGrid20.Name = "lblGrid20"
        Me.lblGrid20.Size = New System.Drawing.Size(24, 25)
        Me.lblGrid20.TabIndex = 15
        Me.lblGrid20.Text = "X"
        Me.lblGrid20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid12
        '
        Me.lblGrid12.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid12.Location = New System.Drawing.Point(128, 64)
        Me.lblGrid12.Name = "lblGrid12"
        Me.lblGrid12.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid12.TabIndex = 14
        Me.lblGrid12.Text = "X"
        Me.lblGrid12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid11
        '
        Me.lblGrid11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid11.Location = New System.Drawing.Point(88, 64)
        Me.lblGrid11.Name = "lblGrid11"
        Me.lblGrid11.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid11.TabIndex = 13
        Me.lblGrid11.Text = "X"
        Me.lblGrid11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid10
        '
        Me.lblGrid10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid10.Location = New System.Drawing.Point(48, 64)
        Me.lblGrid10.Name = "lblGrid10"
        Me.lblGrid10.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid10.TabIndex = 12
        Me.lblGrid10.Text = "X"
        Me.lblGrid10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid02
        '
        Me.lblGrid02.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid02.Location = New System.Drawing.Point(128, 24)
        Me.lblGrid02.Name = "lblGrid02"
        Me.lblGrid02.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid02.TabIndex = 11
        Me.lblGrid02.Text = "X"
        Me.lblGrid02.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid01
        '
        Me.lblGrid01.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid01.Location = New System.Drawing.Point(88, 24)
        Me.lblGrid01.Name = "lblGrid01"
        Me.lblGrid01.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid01.TabIndex = 10
        Me.lblGrid01.Text = "X"
        Me.lblGrid01.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid00
        '
        Me.lblGrid00.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid00.Location = New System.Drawing.Point(49, 24)
        Me.lblGrid00.Name = "lblGrid00"
        Me.lblGrid00.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid00.TabIndex = 9
        Me.lblGrid00.Text = "X"
        Me.lblGrid00.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCurrentPlayer
        '
        Me.lblCurrentPlayer.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentPlayer.Location = New System.Drawing.Point(412, 384)
        Me.lblCurrentPlayer.Name = "lblCurrentPlayer"
        Me.lblCurrentPlayer.Size = New System.Drawing.Size(399, 23)
        Me.lblCurrentPlayer.TabIndex = 10
        Me.lblCurrentPlayer.Text = "Label1"
        '
        'lblYw
        '
        Me.lblYw.Location = New System.Drawing.Point(825, 485)
        Me.lblYw.Name = "lblYw"
        Me.lblYw.Size = New System.Drawing.Size(56, 23)
        Me.lblYw.TabIndex = 11
        Me.lblYw.Text = "Your Wins"
        '
        'lblYourWins
        '
        Me.lblYourWins.Location = New System.Drawing.Point(889, 485)
        Me.lblYourWins.Name = "lblYourWins"
        Me.lblYourWins.Size = New System.Drawing.Size(40, 23)
        Me.lblYourWins.TabIndex = 12
        '
        'lblMw
        '
        Me.lblMw.Location = New System.Drawing.Point(825, 517)
        Me.lblMw.Name = "lblMw"
        Me.lblMw.Size = New System.Drawing.Size(56, 23)
        Me.lblMw.TabIndex = 13
        Me.lblMw.Text = "My Wins"
        '
        'lblMyWins
        '
        Me.lblMyWins.Location = New System.Drawing.Point(889, 517)
        Me.lblMyWins.Name = "lblMyWins"
        Me.lblMyWins.Size = New System.Drawing.Size(40, 23)
        Me.lblMyWins.TabIndex = 14
        '
        'lblD
        '
        Me.lblD.Location = New System.Drawing.Point(827, 549)
        Me.lblD.Name = "lblD"
        Me.lblD.Size = New System.Drawing.Size(56, 23)
        Me.lblD.TabIndex = 15
        Me.lblD.Text = "Draws"
        '
        'lblDraws
        '
        Me.lblDraws.Location = New System.Drawing.Point(889, 549)
        Me.lblDraws.Name = "lblDraws"
        Me.lblDraws.Size = New System.Drawing.Size(40, 23)
        Me.lblDraws.TabIndex = 16
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2})
        Me.MenuItem1.Text = "&Configuración"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "&AI"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(704, -82)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(45, 13)
        Me.Label10.TabIndex = 48
        Me.Label10.Text = "Label10"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(412, 207)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 36)
        Me.Button1.TabIndex = 44
        Me.Button1.Text = "JUGAR ROBOT"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'pctJuego
        '
        Me.pctJuego.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pctJuego.ErrorImage = Nothing
        Me.pctJuego.Location = New System.Drawing.Point(12, 430)
        Me.pctJuego.Name = "pctJuego"
        Me.pctJuego.Size = New System.Drawing.Size(360, 360)
        Me.pctJuego.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pctJuego.TabIndex = 43
        Me.pctJuego.TabStop = False
        '
        'TimMuestreo
        '
        Me.TimMuestreo.Interval = 200
        '
        'PicResultado
        '
        Me.PicResultado.ErrorImage = Nothing
        Me.PicResultado.Location = New System.Drawing.Point(15, 47)
        Me.PicResultado.Name = "PicResultado"
        Me.PicResultado.Size = New System.Drawing.Size(360, 360)
        Me.PicResultado.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PicResultado.TabIndex = 42
        Me.PicResultado.TabStop = False
        '
        'WebCam1
        '
        Me.WebCam1.AutoSize = True
        Me.WebCam1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.WebCam1.Imagen = Nothing
        Me.WebCam1.Location = New System.Drawing.Point(998, 12)
        Me.WebCam1.MilisegundosCaptura = 100
        Me.WebCam1.Name = "WebCam1"
        Me.WebCam1.Size = New System.Drawing.Size(360, 360)
        Me.WebCam1.TabIndex = 49
        Me.WebCam1.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(499, 268)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 50
        Me.Label3.Text = "Label3"
        Me.Label3.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(454, 302)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 51
        Me.Label5.Text = "Label5"
        Me.Label5.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(499, 302)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 13)
        Me.Label6.TabIndex = 52
        Me.Label6.Text = "Label6"
        Me.Label6.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(409, 337)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(39, 13)
        Me.Label7.TabIndex = 53
        Me.Label7.Text = "Label7"
        Me.Label7.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(454, 337)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(39, 13)
        Me.Label8.TabIndex = 54
        Me.Label8.Text = "Label8"
        Me.Label8.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(499, 337)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 55
        Me.Label9.Text = "Label9"
        Me.Label9.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(409, 268)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 56
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(454, 268)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "Label2"
        Me.Label2.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(409, 302)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 13)
        Me.Label4.TabIndex = 58
        Me.Label4.Text = "Label4"
        Me.Label4.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label11.Location = New System.Drawing.Point(12, 12)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(81, 13)
        Me.Label11.TabIndex = 59
        Me.Label11.Text = "Imagen Cámara"
        '
        'Serial
        '
        Me.Serial.PortName = "COM3"
        '
        'frmRobot
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1370, 728)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.WebCam1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.pctJuego)
        Me.Controls.Add(Me.PicResultado)
        Me.Controls.Add(Me.lblDraws)
        Me.Controls.Add(Me.lblD)
        Me.Controls.Add(Me.lblMyWins)
        Me.Controls.Add(Me.lblMw)
        Me.Controls.Add(Me.lblYourWins)
        Me.Controls.Add(Me.lblYw)
        Me.Controls.Add(Me.lblCurrentPlayer)
        Me.Controls.Add(Me.GroupBox1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmRobot"
        Me.Text = "Robot playing Tic Tac Toe SistemasyMicros"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.pctJuego, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicResultado, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Dim oPen As Pen
    Dim oGrafico As Graphics
    Dim modeloRojo As Byte = 75

    Dim largoCuadroX As Byte = 60
    Dim largoCuadroY As Byte = 80
    Dim cantidadMuestra As Byte = 220
    Dim paso As Byte = 5
    'Subrutina del botón de toma de muestra


    'Subrutina del botón Detener
    Private Sub ButDetener_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        WebCam1.Stop()  'Simplemente detiene la camara y deja los últimos fotogramas
        TimMuestreo.Enabled = False
    End Sub

    'Subrutina del botón de configuración
    Private Sub ButConfiguracion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        WebCam1.Configuracion() 'Abre una pantalla de configuracion de la webcam
    End Sub

    'Subrutina del botón de Búsqueda de Azul
    Private Sub ButBuscaAzul_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TimMuestreo_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimMuestreo.Tick

        'Pasamos el fotograma actual al PictureBox para ver lo que hay
        PicResultado.Image = WebCam1.Imagen

        'Declaramos un par de objetos Bitmap para trabajar sin que se vea.
        'Declararemos dos para facilitar la comprensión del código
        Dim ImgOriginal As Bitmap = PicResultado.Image

        Dim ImgResultante As Bitmap = PicResultado.Image

        'Declaramos una variable para cada color. Son de tipo entero porque 
        'Serán valores entre 0 y 255
        Dim Rojo, Verde, Azul As Integer

        'La tolerancia nos permitira no ser tan estrictos con la selección de color
        Dim Tolerancia As Integer

        'Las variables x e y serviran para guardar coordenadas actuales
        Dim x, y As Integer

        'Pondremos una tolerancia de 5 unidades de color. cada uno 
        'debe ajustarla según sus necesidades
        Tolerancia = 0

        'Ahora recorreremos la imagen del objeto ImgOriginal como si fuese una 
        'matriz, de modo que iremos recorriendo las horizontales en orden.
        
        For x = 60 To 240 - 1 Step paso

            For y = 120 To ImgOriginal.Height - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                Rojo = ImgOriginal.GetPixel(x, y).R
                Verde = ImgOriginal.GetPixel(x, y).G
                Azul = ImgOriginal.GetPixel(x, y).B


                'Ahora preguntaremos si en el pixel actual predomina el 
                'Azul sobre el rojo y el verde, teniendo en cuenta la tolerancia
                'If Rojo >= Azul - Tolerancia Then   'Si hay más azul que rojo sigue
                If Rojo >= Verde - Tolerancia Then  'Si hay mas azul que verde sigue

                    'si llega hasta aqui es porque el pixel es suficientemente azul
                    'asi que lo pintamos con un punto de 2x2 en amarillo
                    'esta es una manera de pintar cuadrados, pero luego vemos otra

                    ImgResultante.SetPixel(x, y, Color.Red)
                    

                End If
            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)


        PicResultado.Image = ImgResultante

    End Sub

    Private Sub decidir()
        Dim x, y As Integer
        Dim cantidadRojo As Integer
        Dim rojo As Color
        Dim imgJuego As Bitmap = PicResultado.Image
        'Dim imgJuegoResultante As Bitmap
        '---------------------------------------------------------------------------------
        'Reviso el primer cuadro
        For x = 60 To largoCuadroX + 60 - 1 Step paso


            For y = 120 To largoCuadroY + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)


                If rojo.R >= cantidadMuestra Then

                    cantidadRojo = cantidadRojo + 1

                End If

            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        'Label1.Text = cantidadRojo

        oGrafico = pctJuego.CreateGraphics

        Dim args As System.EventArgs


        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid00, args)

        Else
            oPen = New Pen(Color.Green, 1)


        End If
        oGrafico.DrawArc(oPen, 50, 50, 20, 20, 0, 360)
        '-------------------------------------------------------------------------------

        'Reviso el segundo cuadro

        cantidadRojo = 0

        For x = largoCuadroX + 60 To largoCuadroX * 2 + 60 - 1 Step paso

            For y = 120 To largoCuadroY + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)

                If rojo.R >= cantidadMuestra Then


                    cantidadRojo = cantidadRojo + 1

                End If

            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        'Label2.Text = cantidadRojo

        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid01, args)

        Else
            oPen = New Pen(Color.Green, 1)


        End If

        oGrafico.DrawArc(oPen, 100, 50, 20, 20, 0, 360)


        '---------------------------------------------------------------------------
        'Reviso el tercer cuadro

        cantidadRojo = 0

        For x = largoCuadroX * 2 + 60 To largoCuadroX * 3 + 60 - 1 Step paso

            For y = 120 To largoCuadroY + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)

                If rojo.R >= cantidadMuestra Then

                    cantidadRojo = cantidadRojo + 1

                End If

            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        'Label3.Text = cantidadRojo

        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid02, args)

        Else

            oPen = New Pen(Color.Green, 1)

        End If

        oGrafico.DrawArc(oPen, 150, 50, 20, 20, 0, 360)

        '--------------------------------------------------------------------------------

        'Reviso el cuarto cuadro

        cantidadRojo = 0

        For x = 60 To largoCuadroX + 60 - 1 Step paso

            For y = largoCuadroY + 120 To largoCuadroY * 2 + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)
                If rojo.R > cantidadMuestra Then
                    cantidadRojo = cantidadRojo + 1
                End If
            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        'Label4.Text = cantidadRojo

        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid10, args)

        Else

            oPen = New Pen(Color.Green, 1)

        End If

        oGrafico.DrawArc(oPen, 50, 100, 20, 20, 0, 360)

        '--------------------------------------------------------------------------------

        'Reviso el quinto cuadro

        cantidadRojo = 0

        For x = largoCuadroX + 60 To largoCuadroX * 2 + 60 - 1 Step paso

            For y = largoCuadroY + 120 To largoCuadroY * 2 + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)

                If rojo.R > cantidadMuestra Then
                    cantidadRojo = cantidadRojo + 1
                End If
            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        '        Label5.Text = cantidadRojo

        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid11, args)

        Else

            oPen = New Pen(Color.Green, 1)


        End If

        oGrafico.DrawArc(oPen, 100, 100, 20, 20, 0, 360)

        '--------------------------------------------------------------------------------
        'Reviso el sexto cuadro

        cantidadRojo = 0

        For x = largoCuadroX * 2 + 60 To largoCuadroX * 3 + 60 - 1 Step paso

            For y = largoCuadroY + 120 To largoCuadroY * 2 + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)
                If rojo.R > cantidadMuestra Then
                    cantidadRojo = cantidadRojo + 1
                End If
            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        '        Label6.Text = cantidadRojo

        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid12, args)

        Else

            oPen = New Pen(Color.Green, 1)


        End If


        oGrafico.DrawArc(oPen, 150, 100, 20, 20, 0, 360)

        '--------------------------------------------------------------------------------

        'Reviso el septimo cuadro

        cantidadRojo = 0

        For x = 60 To largoCuadroX + 60 - 1 Step paso

            For y = largoCuadroY * 2 + 120 To largoCuadroY * 3 + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)

                If rojo.R > cantidadMuestra Then

                    cantidadRojo = cantidadRojo + 1

                End If

            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        '        Label7.Text = cantidadRojo

        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid20, args)

        Else

            oPen = New Pen(Color.Green, 1)


        End If


        oGrafico.DrawArc(oPen, 50, 150, 20, 20, 0, 360)

        '--------------------------------------------------------------------------------

        'Reviso el octavo cuadro

        cantidadRojo = 0

        For x = largoCuadroX + 60 To largoCuadroX * 2 + 60 - 1 Step paso

            For y = largoCuadroY * 2 + 120 To largoCuadroY * 3 + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)

                If rojo.R > cantidadMuestra Then

                    cantidadRojo = cantidadRojo + 1

                End If

            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        'Label8.Text = cantidadRojo

        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid21, args)

        Else

            oPen = New Pen(Color.Green, 1)


        End If


        oGrafico.DrawArc(oPen, 100, 150, 20, 20, 0, 360)


        '--------------------------------------------------------------------------------

        'Reviso el noveno cuadro

        cantidadRojo = 0

        For x = largoCuadroX * 2 + 60 To largoCuadroX * 3 + 60 - 1 Step paso

            For y = largoCuadroY * 2 + 120 To largoCuadroY * 3 + 120 - 1 Step paso

                'Guardamos los valores de los colores del pixel (x, y)
                rojo = imgJuego.GetPixel(x, y)

                If rojo.R > cantidadMuestra Then

                    cantidadRojo = cantidadRojo + 1

                End If

            Next    'Fin del analisis del pixel x, y

        Next    'Fin de la linea horizontal pasa a la siguiente (y+1)

        'Label9.Text = cantidadRojo

        If cantidadRojo >= modeloRojo Then

            ' para la linea ( color rojo con 1 de grosor )  
            oPen = New Pen(Color.Red, 1)

            lblGrid_Click(lblGrid22, args)

        Else

            oPen = New Pen(Color.Green, 1)


        End If


        oGrafico.DrawArc(oPen, 150, 150, 20, 20, 0, 360)



        '--------------------------------------------------------------------------------

    End Sub

    Private Sub PicResultado_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicResultado.MouseMove
        Label10.Text = e.X & " " & e.Y
    End Sub
    'Grid Code
    Public Sub lblGrid_Click(ByVal sender As Object, ByVal args As EventArgs)
        Try
            Dim aLabelControl As Label = CType(sender, Label)
            Dim sName As String = aLabelControl.Name
            Dim iRow As Integer = sName.Substring(sName.Length - 2, 1)
            Dim iCol As Integer = sName.Substring(sName.Length - 1, 1)

            If m_TicTacToeGame.Square(iRow, iCol) = TicTacToe.GridEntry.NoEntry Then
                m_TicTacToeGame.Square(iRow, iCol) = m_iCurrentPlayer
                ToggleCurrentPlayer()
                DrawGrid()
            End If
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub ToggleCurrentPlayer()
        If m_iCurrentPlayer = TicTacToe.GridEntry.PlayerX Then
            m_iCurrentPlayer = TicTacToe.GridEntry.PlayerO
            If m_bIsAIEnabled = True And m_bGameEnded = False Then
                PerformAIMove()
            End If
        Else
            m_iCurrentPlayer = TicTacToe.GridEntry.PlayerX
        End If
    End Sub

    Private Sub DrawGrid()
        Dim iRow As Integer
        Dim iCol As Integer
        Dim aLabelControl As Label

        For iRow = 0 To 2
            For iCol = 0 To 2
                Try
                    aLabelControl = GetLabelControl(iRow, iCol)
                    Select Case m_TicTacToeGame.Square(iRow, iCol)
                        Case TicTacToe.GridEntry.NoEntry
                            aLabelControl.Text = ""
                        Case TicTacToe.GridEntry.PlayerX
                            aLabelControl.Text = "X"
                        Case TicTacToe.GridEntry.PlayerO

                            aLabelControl.Text = "O"
                            'MsgBox(iRow)
                            'MsgBox(iCol)
                    End Select
                Catch ex As Exception

                End Try
            Next
        Next

        updatelabels()
    End Sub

    Private Sub PerformAIMove()
        'Aca el robot decide donde mover la ficha
        Try
            Dim aSC As SquareCoordinate
            Try
                aSC = m_TicTacToeAI.GetAIMove(m_TicTacToeGame)

                posicionFicha = aSC.Row & aSC.Column


                Select Case posicionFicha

                    Case "00"

                        Serial.Write("r")
                        Serial.Write("9")

                    Case "01"

                        Serial.Write("r")
                        Serial.Write("8")

                    Case "02"

                        Serial.Write("r")
                        Serial.Write("7")

                    Case "10"

                        Serial.Write("r")
                        Serial.Write("6")

                    Case "11"

                        Serial.Write("r")
                        Serial.Write("5")

                    Case "12"

                        Serial.Write("r")
                        Serial.Write("4")

                    Case "20"

                        Serial.Write("r")
                        Serial.Write("3")

                    Case "21"

                        Serial.Write("r")
                        Serial.Write("2")

                    Case "22"

                        Serial.Write("r")
                        Serial.Write("1")


                End Select

            Catch ex As NullReferenceException
                Stop
            Catch ex As Exception
                Stop
            End Try
            m_TicTacToeGame.Square(aSC) = TicTacToe.GridEntry.PlayerO
            DrawGrid()
            ToggleCurrentPlayer()
        Catch ex As Exception
            Stop
        End Try

    End Sub

    Private Function GetLabelControl(ByVal iRow As Integer, ByVal iCol As Integer) As Label
        Dim x As Integer
        Dim aControl As Control
        For x = 0 To GroupBox1.Controls.Count - 1
            aControl = GroupBox1.Controls(x)
            If aControl.Name = "lblGrid" & iRow & iCol Then Return aControl
        Next
    End Function

    Private Sub Cat()
        DrawGrid()
        MessageBox.Show("A Cat Has Occured ... There is no winners.")
        m_iDrawCount = m_iDrawCount + 1
        m_bGameEnded = True
        NewGame()
    End Sub

    Private Sub WeHaveAWinner(ByVal iWinner As TicTacToe.GridEntry)
        DrawGrid()
        Dim sMessage As String = "Player "
        Select Case iWinner
            Case TicTacToe.GridEntry.PlayerX
                sMessage = sMessage & " X has won."
                m_iPlayerXCount = m_iPlayerXCount + 1
            Case TicTacToe.GridEntry.PlayerO
                sMessage = sMessage & " O has won."
                Serial.Write("w")
                m_iPlayerOCount = m_iPlayerOCount + 1
        End Select
        'MessageBox.Show(sMessage)
        'm_bGameEnded = True
        'NewGame()
    End Sub

    Private Sub NewGame()
        Dim rtnDlgResult As DialogResult
        rtnDlgResult = MessageBox.Show("Would you like to play another game?", "New Game", MessageBoxButtons.YesNo)
        If rtnDlgResult = DialogResult.Yes Then
            m_TicTacToeGame.Clear()
            m_bGameEnded = False
        Else
            Me.Close()
        End If
    End Sub

    Private Sub updatelabels()
        Select Case m_iCurrentPlayer
            Case TicTacToe.GridEntry.PlayerX
                lblCurrentPlayer.Text = "Turno del jugador"
            Case TicTacToe.GridEntry.PlayerO
                lblCurrentPlayer.Text = "Turno del robot"
        End Select

        lblYourWins.Text = m_iPlayerXCount
        lblMyWins.Text = m_iPlayerOCount
        lblDraws.Text = m_iDrawCount
    End Sub

    'AI Setting
    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        If m_bIsAIEnabled = True Then
            m_bIsAIEnabled = False
            MenuItem2.Checked = False
            Me.ReSet()
        Else
            m_bIsAIEnabled = True
            Me.ReSet()
            MenuItem2.Checked = True
        End If
    End Sub

    Private Sub ReSet()
        m_TicTacToeGame.Clear()
        m_iPlayerXCount = 0
        m_iPlayerOCount = 0
        m_iDrawCount = 0
        DrawGrid()
    End Sub

    Private Sub ActivarCamara()
        'On Error Resume Next
        WebCam1.Start() 'Activa la WebCam

        PicResultado.Image = WebCam1.Imagen 'Pasa una muestra al PictureBox para ver el resultado

    End Sub


    Private Sub btnBuscaRojo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'WebCam1.Start() 'Activa la WebCam
        TimMuestreo.Enabled = True   'Activa el timer que nos generará el muestreo
    End Sub


    Private Sub frmRobot_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Serial.Open()   'Abro el puerto serial

        ActivarCamara()

        TimMuestreo.Enabled = True   'Activa el timer que nos generará el muestreo

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        decidir()

    End Sub

    Private Sub lblYw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblYw.Click
        Serial.Write("i")
    End Sub

    Private Sub lblD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblD.Click
        m_bGameEnded = True
        NewGame()
    End Sub
End Class
