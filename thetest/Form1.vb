Public Class Form1
    Dim mSettingsDict As Dictionary(Of String, Int32)

    Shared Sub Main()
        If Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MessageBox.Show("多重起動はできません。")
            Return
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TB_CurrentDir.Text = IO.Directory.GetCurrentDirectory()
        Dim filename As String = "settings.txt"
        If IO.File.Exists(IO.Directory.GetCurrentDirectory() & "\" & filename) = True Then
            Dim sr As IO.StreamReader = New IO.StreamReader(TB_CurrentDir.Text & "\" & filename)
            Dim str As String = ""
            Do While sr.EndOfStream = False
                str = sr.ReadLine
                '                TB_ML_Settings.AppendText(str & vbCrLf)
            Loop

            '            TB_ML_Settings.Text = str
            sr.Close()
        End If
    End Sub

    Private Sub 設定ToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FolderBrowserDialog1 As New FolderBrowserDialog()

        FolderBrowserDialog1.Description = "データの出力フォルダを指定"

        If TB_OutputDirectory.TextLength < 1 Then
            FolderBrowserDialog1.SelectedPath = System.IO.Directory.GetCurrentDirectory()
        Else
            FolderBrowserDialog1.SelectedPath = TB_OutputDirectory.Text
        End If

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            MessageBox.Show(FolderBrowserDialog1.SelectedPath)
            TB_OutputDirectory.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Btn_opencsvfiledialog_Click(sender As Object, e As EventArgs) Handles Btn_opencsvfiledialog.Click
        Dim opencsvfiledlg As OpenFileDialog = New OpenFileDialog
        opencsvfiledlg.Filter = "CSVファイル(*.csv)|*.csv"
        opencsvfiledlg.FilterIndex = 1
        If opencsvfiledlg.ShowDialog() = DialogResult.OK Then
            TB_CSVFilePath.Text = opencsvfiledlg.FileName
        End If


    End Sub

    Dim mListOriginalString As List(Of String) = New List(Of String)
    Dim mDictJumpLabel As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If String.IsNullOrWhiteSpace(TB_CSVFilePath.Text) = True Then
            MessageBox.Show("ファイルパスが正しく設定されていません")
            Return
        End If
        If IO.File.Exists(TB_CSVFilePath.Text) = False Then
            MessageBox.Show("ファイルが存在していません")
            Return
        End If
        ComboBox1.Items.Clear()
        Dim sr As IO.StreamReader = New IO.StreamReader(TB_CSVFilePath.Text)
        Dim str As String = ""
        Dim index As Int32 = New Int32
        index = 1
        Do While sr.EndOfStream = False
            str = sr.ReadLine
            'コメント行、カンマが１つもみつからなかったらパス
            If str.IndexOf("//") = 0 Or str.IndexOf(",") = -1 Then
                Continue Do
            End If
            '空行パス
            If String.IsNullOrWhiteSpace(str) Then
                Continue Do
            End If


            Dim strarray() As String = str.Split(",")
            If strarray.Length < 3 Then
                MessageBox.Show(index & "要素目のフォーマットがおかしいです：要素３未満")
                Continue Do
            End If
            '第三要素がハイフンならばパス
            If strarray.GetValue(2) = "-" Then
                Continue Do
            End If
            '１番目と２番目の要素が空だと無効
            If String.IsNullOrEmpty(strarray.GetValue(0)) Or String.IsNullOrEmpty(strarray.GetValue(1)) Then
                MessageBox.Show(index & "要素目のフォーマットがおかしいです：要素1&2が未設定")
                Continue Do
            End If

            mListOriginalString.Add(str)

            TB_ML_Settings.AppendText(strarray.Length & " " & str & vbCrLf)
            index = index + 1
            ComboBox1.Items.Add(strarray(1))
        Loop
        sr.Close()

    End Sub

    Private Function GetIDFromDescription(desc As String) As Integer
        For Each str As String In mListOriginalString
            If str.IndexOf(desc) > 0 Then
                Dim array As String() = str.Split(",")
                Return Integer.Parse(array.GetValue(0))
            End If
        Next
        Return -1
    End Function

    Private Function GetStrFromDescription(desc As String) As String
        For Each str As String In mListOriginalString
            If str.IndexOf(desc) > 0 Then
                Return str
            End If
        Next
        Return ""
    End Function


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem IsNot Nothing Then
            Dim id As Integer = GetIDFromDescription(ComboBox1.SelectedItem)
            If id <> -1 Then
                Label_CodeID.Text = "コードID:" & id
                Dim str As String = GetStrFromDescription(ComboBox1.SelectedItem)
                Dim strarray As String() = str.Split(",")
                Dim description = strarray.GetValue(1)
                Dim tmplist As List(Of String) = New List(Of String)
                Dim i As Integer = -1
                Dim valuecount As Integer = 0
                For Each tmpstr As String In strarray
                    i = i + 1
                    If i = 0 Or i = 1 Then
                        Continue For
                    End If
                    If i = 2 Then
                        Try
                            valuecount = Integer.Parse(tmpstr)
                        Catch err As FormatException


                        End Try
                    End If
                    tmplist.Add(tmpstr)
                Next
                'Dim opcode As Integer = Integer.Parse(tmplist(i))
                Select Case tmplist(2)
                    Case "0"
                        TabControl2.SelectedTab = TabPage_Int
                    Case "P"
                        TabControl2.SelectedTab = TabPage_Int
                    Case "J"
                        TabControl2.SelectedTab = TabPage_Jump
                    Case "L"
                        TabControl2.SelectedTab = TabPage_Label

                End Select
            End If
        End If
    End Sub

End Class
