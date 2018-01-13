Imports System.Speech.Recognition
Imports System.Threading
Imports System.Globalization
Public Class Form1
    ' recogniser & grammar
    Dim recog As New SpeechRecognizer
    Dim gram As Grammar
    ' events
    Public Event SpeechRecognized As _
        EventHandler(Of SpeechRecognizedEventArgs)
    Public Event SpeechRecognitionRejected As _
        EventHandler(Of SpeechRecognitionRejectedEventArgs)
    ' word list
    Dim wordlist As String() = New String() {"Yes", "No", "Maybe"}
    ' word recognised event
    Public Sub recevent(ByVal sender As System.Object,
            ByVal e As RecognitionEventArgs)
        LabelYes.ForeColor = Color.LightGray
        LabelNo.ForeColor = Color.LightGray
        LabelMaybe.ForeColor = Color.LightGray
        If (e.Result.Text = "Yes") Then
            LabelYes.ForeColor = Color.Blue
        ElseIf (e.Result.Text = "No") Then
            LabelNo.ForeColor = Color.Blue
        ElseIf (e.Result.Text = "Maybe") Then
            LabelMaybe.ForeColor = Color.Blue
        End If
    End Sub
    ' recognition failed event
    Public Sub recfailevent(ByVal sender As System.Object,
            ByVal e As RecognitionEventArgs)
        LabelYes.ForeColor = Color.LightGray
        LabelNo.ForeColor = Color.LightGray
        LabelMaybe.ForeColor = Color.LightGray
    End Sub
    ' form initialisation
    Private Sub Form1_Load(ByVal sender As System.Object,
            ByVal e As System.EventArgs) Handles MyBase.Load
        ' need these to get British English rather than default US
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-GB")
        ' convert the word list into a grammar
        Dim words As New Choices(wordlist)
        gram = New Grammar(New GrammarBuilder(words))
        recog.LoadGrammar(gram)
        ' add handlers for the recognition events
        AddHandler recog.SpeechRecognized, AddressOf Me.recevent
        AddHandler recog.SpeechRecognitionRejected, AddressOf Me.recfailevent
        ' enable the recogniser
        recog.Enabled = True
    End Sub
End Class