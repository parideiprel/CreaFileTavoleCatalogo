Imports System.IO

Public Class Form1





    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Definizione variabili
        Dim iIdx As Integer = 0
        Dim fsIn As FileStream
        Dim binStreamReader As BinaryReader
        Dim binIn As Array

        Dim fsOut As FileStream

        Dim sPre As String
        Dim iStart As Integer
        Dim iEnd As Integer
        Dim sPost As String

        Dim sTavola As String
        Select Case TextBox1.Text
            Case Is = ""
                MsgBox("Inserire codice quadro!", vbExclamation & vbOKOnly, "Dati mancanti")
                If TextBox1.Text = "" Then
                    TextBox1.BackColor = Color.Red
                End If
                TextBox1.Focus()
                Exit Sub
            Case Is <> ""
                TextBox1.BackColor = Color.White
                TextBox1.Text = UCase(TextBox1.Text)
            Case Else
                Exit Select
        End Select

        Select Case TextBox2.Text
            Case Is = ""
                MsgBox("Inserire estensione!", vbExclamation & vbOKOnly, "Dati mancanti")
                If TextBox2.Text = "" Then
                    TextBox2.BackColor = Color.Red
                End If
                TextBox2.Focus()
                Exit Sub
            Case Is <> ""
                TextBox2.BackColor = Color.White
            Case Else
                Exit Select
        End Select



        With ProgressBar1
            .Minimum = iStart
            .Maximum = iEnd
            .ForeColor = Color.Blue
            .Step = 1
        End With

        Try
            fsIn = New FileStream(Application.StartupPath & "\Matrici\Tavola_Bianca.psd", FileMode.Open)

            binStreamReader = New BinaryReader(fsIn)
            binIn = binStreamReader.ReadBytes(Convert.ToInt32(fsIn.Length))
            binStreamReader.Close()
            fsIn.Close()

            Dim a = FolderBrowserDialog1.ShowDialog()
            If a = Windows.Forms.DialogResult.Cancel Then
                MsgBox("Generazione annullata", vbOKOnly, "Intervento Utente")
                Exit Sub
            End If

            'compongo il nome del file
            sPre = TextBox1.Text
            iStart = NumericUpDown1.Value
            iEnd = NumericUpDown2.Value
            sPost = TextBox2.Text

            For iIdx = iStart To iEnd

                ProgressBar1.PerformStep()

                'If iEnd > 9 Then
                If iEnd > 9 And iIdx < 10 Then
                    sTavola = sPre & "A" & "0" & iIdx.ToString & iEnd.ToString & "IJ" & sPost & ".psd"
                Else
                    sTavola = sPre & "A" & iIdx.ToString & iEnd.ToString & "IJ" & sPost & ".psd"
                End If


                fsOut = New FileStream(FolderBrowserDialog1.SelectedPath & "\" & sTavola, FileMode.Create)
                fsOut.Write(binIn, 0, binIn.Length)
                fsOut.Close()

            Next



        Catch ex As Exception
            MsgBox("Exception: " & ex.Message)
            
            If fsIn.SafeFileHandle.IsClosed = False Then
                fsIn.Close()
                fsIn.Dispose()
            Else
                fsIn.Dispose()
            End If

            If fsOut.SafeFileHandle.IsClosed = False Then
                fsOut.Close()
                fsOut.Dispose()
            Else
                fsOut.Dispose()
            End If

            Exit Sub


        End Try

        MsgBox("Generazione Terminata !!", vbExclamation & vbOKOnly, "Esito Generazione")
        '        If fsIn.SafeFileHandle.IsInvalid = False Then
        'fsIn.Close()
        fsIn.Dispose()
        'Else
        'fsIn.Dispose()
        'End If

        'If fsOut.SafeFileHandle.IsClosed = False Then
        'fsOut.Close()
        fsOut.Dispose()
        'Else
        'fsOut.Dispose()
        'End If

        '        fsIn.Dispose()
        '        fsOut.Dispose()
        End

        'Function LoadBinaryData(ByVal path As String) As Object
        '        ' Open a file stream for input
        '        Dim fs As FileStream = New FileStream(path, FileMode.Open)
        '        ' Create a binary formatter for this stream
        '        Dim bf As New BinaryFormatter()

        '        ' Deserialize the contents of the file stream into an object
        '        LoadBinaryData = bf.Deserialize(fs)
        '        ' close the stream.
        '        fs.Close()
        '    End Function

        '        FileStream fs = new FileStream(filePath, FileMode.Open);
        'BinaryReader br = new BinaryReader(fs);
        'byte[] bin = br.ReadBytes(Convert.ToInt32(fs.Length));
        'fs.Close();
        'br.Close();



    End Sub
End Class
