Public Class Form1
    Dim mIntLabelControls As Label()
    Dim marTextBox_Int As TextBox()
    Dim mDictOrderDescriptionAndID As Dictionary(Of String, Integer)
    Dim mDictOrderIDAndDescription As Dictionary(Of Integer, String)
    Dim mstrLabelAddrEscapeSequence As String = "$LBaddr"

    Shared Sub Main()
        If Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MessageBox.Show("多重起動はできません。")
            Return
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If IO.File.Exists(My.Settings.OutputFilePath) = True Then
            TB_CSVFilePath.Text = My.Settings.OutputFilePath
            ParseDefinitionFile()
        Else
        End If
        TB_CurrentDir.Text = IO.Directory.GetCurrentDirectory()
        Dim filename As String = "settings.txt"
        If IO.File.Exists(IO.Directory.GetCurrentDirectory() & "\" & filename) = True Then
            Dim sr As IO.StreamReader = New IO.StreamReader(TB_CurrentDir.Text & "\" & filename)
            Dim str As String = ""
            Do While sr.EndOfStream = False
                str = sr.ReadLine
                'TB_ML_Settings.AppendText(str & vbCrLf)
            Loop

            'TB_ML_Settings.Text = str
            sr.Close()
        End If
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox_JumpLabel.DropDownStyle = ComboBoxStyle.DropDownList
        HideAllTab()
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
            My.Settings.OutputFilePath = opencsvfiledlg.FileName
        End If
    End Sub

    Dim mListOriginalString As List(Of String) = New List(Of String)
    Dim mDictJumpLabel As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
    Dim miLabelCounter As Integer = 0
    Dim miJumpOrderID As Integer = -1
    Dim miLabelOrderID As Integer = -1
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ParseDefinitionFile()
    End Sub

    Private Sub ParseDefinitionFile()
        If String.IsNullOrWhiteSpace(TB_CSVFilePath.Text) = True Then
            MessageBox.Show("ファイルパスが正しく設定されていません")
            Return
        End If
        If IO.File.Exists(TB_CSVFilePath.Text) = False Then
            MessageBox.Show("ファイルが存在していません")
            Return
        End If
        ComboBox1.Items.Clear()
        mDictOrderDescriptionAndID = New Dictionary(Of String, Integer)
        mDictOrderIDAndDescription = New Dictionary(Of Integer, String)
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
            '3番目がハイフンならばパス
            If strarray.GetValue(2) = "-" Then
                Continue Do
            End If
            '0番目と1番目の要素が空だと無効
            If String.IsNullOrEmpty(strarray.GetValue(0)) Or String.IsNullOrEmpty(strarray.GetValue(1)) Then
                MessageBox.Show(index & "要素目のフォーマットがおかしいです：要素1&2が未設定")
                Continue Do
            End If

            'ジャンプコードを記憶
            If strarray.GetValue(2) = "J" Then
                miJumpOrderID = Integer.Parse(strarray.GetValue(0))
            End If
            'ラベル設定コードを記憶
            If strarray.GetValue(2) = "L" Then
                miLabelOrderID = Integer.Parse(strarray.GetValue(0))
            End If

            mListOriginalString.Add(str)
            TB_ML_Settings.AppendText(strarray.Length & " " & str & vbCrLf)
            index = index + 1

            If mDictOrderDescriptionAndID.ContainsKey(strarray(1)) = True Then
                MessageBox.Show(strarray(1) & ":同じ命令詳細が登録されています")
            Else
                mDictOrderDescriptionAndID.Add(strarray(1), Integer.Parse(strarray(0)))
                mDictOrderIDAndDescription.Add(Integer.Parse(strarray(0)), strarray(1))
            End If
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

    Private Sub HideAllTab()
        TabControl2.TabPages.Remove(TabPage_Int)
        TabControl2.TabPages.Remove(TabPage_Compare)
        TabControl2.TabPages.Remove(TabPage_Jump)
        TabControl2.TabPages.Remove(TabPage_Label)
        TabControl2.TabPages.Remove(TabPage_Nothing)
    End Sub

    Private Sub ShowTabIndexOf(index As Integer)
        'インサート先は全部先頭なので0、その代わりhidealltabを事前に使って全消し必須
        Select Case index
            Case 0
                TabControl2.TabPages.Insert(0, TabPage_Int)
            Case 1
                TabControl2.TabPages.Insert(0, TabPage_Compare)
            Case 2
                TabControl2.TabPages.Insert(0, TabPage_Jump)
            Case 3
                TabControl2.TabPages.Insert(0, TabPage_Label)
            Case 4
                TabControl2.TabPages.Insert(0, TabPage_Nothing)
        End Select
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem IsNot Nothing Then
            mIntLabelControls = New Label() {Label_detail1, Label_detail2, Label_detail3, Label_detail4}
            Dim id As Integer = GetIDFromDescription(ComboBox1.SelectedItem)
            If id <> -1 Then
                HideAllTab()
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
                Dim counter As Integer
                Select Case tmplist(2)
                    Case "0"
                        TabControl2.SelectedTab = TabPage_Int
                        For counter = 0 To valuecount
                            Try
                                mIntLabelControls(counter).Text = tmplist(3 + 3 * counter)
                            Catch ex As ArgumentOutOfRangeException
                                MessageBox.Show("定義ファイルに引数の説明がありません")
                                Exit For
                            End Try
                            counter = counter + 1
                        Next
                        ChangeIntTextboxRange(0, valuecount - 1, True)
                        ChangeIntTextboxRange(valuecount, marTextBox_Int.Length - 1, False)
                        ShowTabIndexOf(0)
                    Case "P"
                        TabControl2.SelectedTab = TabPage_Int
                        ShowTabIndexOf(1)
                    Case Else
                        '上のチェックを抜けたら３つあるフォーマット指定の一番左をみる
                        Select Case tmplist(0)
                            Case "J"
                                TabControl2.SelectedTab = TabPage_Jump
                                ShowTabIndexOf(2)
                            Case "L"
                                TabControl2.SelectedTab = TabPage_Label
                                ShowTabIndexOf(3)
                            Case Else 'シートでは無指定、オペランドなしのタブ
                                TabControl2.SelectedTab = TabPage_Nothing
                                ShowTabIndexOf(4)
                        End Select
                End Select
            End If
        End If
    End Sub

    Private Sub ChangeIntTextboxRange(start_index As Integer, end_index As Integer, flag As Boolean)
        marTextBox_Int = New TextBox() {TextBox_Int1, TextBox_Int2, TextBox_Int3, TextBox_Int4}
        For i = start_index To end_index
            marTextBox_Int(i).Enabled = flag
        Next
    End Sub

    Private Sub Button_SetInt_Click(sender As Object, e As EventArgs) Handles Button_SetInt.Click
        If CheckOrderValidation() = False Then
            Return
        End If
        mIntLabelControls = New Label() {Label_detail1, Label_detail2, Label_detail3, Label_detail4}
        marTextBox_Int = New TextBox() {TextBox_Int1, TextBox_Int2, TextBox_Int3, TextBox_Int4}

        Dim tmpstr As String = ComboBox1.SelectedItem
        Dim i As Integer
        For i = 0 To marTextBox_Int.Length - 1
            If mIntLabelControls(i).Text.Length > 0 And String.IsNullOrWhiteSpace(marTextBox_Int(i).Text) = True Then
                MessageBox.Show(i + 1 & "番の入力が不足しています")
                Return
            Else
                '詳細のある項目だけデータとして扱う。詳細なし項目のデータは無視
                'また、入力テキストぼっくすのテキスト長が1以上なければ無視
                If mIntLabelControls(i).Text.Length > 0 And marTextBox_Int.Length > 0 Then
                    If IsNumeric(marTextBox_Int(i).Text) Then
                        tmpstr = tmpstr & "," & marTextBox_Int(i).Text
                    Else
                        MessageBox.Show("入力内容が数値ではありません")
                        Return
                    End If
                End If
            End If

        Next
        AddOrderPack(tmpstr)
    End Sub

    Private Sub Button_ClearInt_Click(sender As Object, e As EventArgs) Handles Button_ClearInt.Click
        marTextBox_Int = New TextBox() {TextBox_Int1, TextBox_Int2, TextBox_Int3, TextBox_Int4}
        For i = 0 To 3
            marTextBox_Int(i).Clear()
        Next
    End Sub

    Private Sub Button_SetCompare_Click(sender As Object, e As EventArgs) Handles Button_SetCompare.Click

    End Sub

    Private Sub Button_ClearCompare_Click(sender As Object, e As EventArgs) Handles Button_ClearCompare.Click

    End Sub

    Private Function IsJumpOrder(orderstring As String) As Boolean
        Dim strarray() As String = orderstring.Split(",")
        If Integer.Parse(strarray.GetValue(0)) = miJumpOrderID Then
            Return True
        End If
        Return False
    End Function

    Private Function IsLabelOrder(orderstring As String) As Boolean
        Dim strarray() As String = orderstring.Split(",")
        If Integer.Parse(strarray.GetValue(0)) = miLabelOrderID Then
            Return True
        End If
        Return False
    End Function

    Private Function GetLabelNameByOrderPack(orderstring As String) As String
        Dim strarray() As String = orderstring.Split(",")
        If IsJumpOrder(orderstring) = True Then
            Return strarray.GetValue(1)
        End If
        Return ""
    End Function

    Private Sub Button_SetJump_Click(sender As Object, e As EventArgs) Handles Button_SetJump.Click
        If CheckOrderValidation() = False Then
            Return
        End If
        If ComboBox_JumpLabel.SelectedIndex <> -1 Then
            Dim tmpstr As String = ComboBox1.SelectedItem & "," & ComboBox_JumpLabel.SelectedItem
            '            Dim iAddr As Integer
            '            If mDictJumpLabel.TryGetValue(ComboBox_JumpLabel.SelectedItem, iAddr) Then
            '            tmpstr = tmpstr & "," & iAddr.ToString()
            AddOrderPack(tmpstr)
            '        End If
        Else
            MessageBox.Show("ジャンプ先ラベルが選択されていません")
        End If
        ComboBox_JumpLabel.SelectedIndex = -1
    End Sub

    Private Sub Button_ClearJump_Click(sender As Object, e As EventArgs) Handles Button_ClearJump.Click
        ComboBox_JumpLabel.SelectedIndex = -1
    End Sub

    Private Sub Button_SetLabelname_Click(sender As Object, e As EventArgs) Handles Button_SetLabelname.Click
        If CheckOrderValidation() = False Then
            Return
        End If
        If String.IsNullOrWhiteSpace(TextBox_JumpLabel.Text) = True Then
            MessageBox.Show("ラベルの文字列には空白・記入無しなどは使えません"）
            Return
        End If

        If mDictJumpLabel.ContainsKey(TextBox_JumpLabel.Text) Then
            MessageBox.Show("すでに使われているラベルです")
            Return
        Else
            mDictJumpLabel.Add(TextBox_JumpLabel.Text, miLabelCounter)
            miLabelCounter = miLabelCounter + 1
        End If
        Dim tmpstr As String = ComboBox1.SelectedItem & "," & TextBox_JumpLabel.Text
        '           & "," & mstrLabelAddrEscapeSequence
        AddOrderPack(tmpstr)
        ComboBox_JumpLabel.Items.Clear()
        For Each str As String In mDictJumpLabel.Keys
            ComboBox_JumpLabel.Items.Add(str)
        Next
        TextBox_JumpLabel.Clear()
    End Sub

    Private Sub Button_ClearLabelname_Click(sender As Object, e As EventArgs) Handles Button_ClearLabelname.Click
        TextBox_JumpLabel.Clear()
    End Sub

    Private Function CheckOrderValidation() As Boolean
        If ComboBox1.SelectedIndex = -1 Then
            MessageBox.Show("命令が選択されていません")
            Return False
        End If
        Return True
    End Function

    Private Sub ListBox_OrderSet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox_OrderSet.SelectedIndexChanged
        If String.IsNullOrEmpty(ListBox_OrderSet.SelectedItem) = False Then
            'TextBox_CurrentOrder.Text = ListBox_OrderSet.SelectedItem
            'MessageBox.Show(ListBox_OrderSet.SelectedItem)
        End If

    End Sub

    Private Sub MoveCursorToLastpos()
        ListBox_OrderSet.SelectedIndex = ListBox_OrderSet.Items.Count - 1
    End Sub

    Private Function IsListboxSelectedValid() As Boolean
        If ListBox_OrderSet.SelectedIndex = -1 Then
            MessageBox.Show("リストボックスの選択がされていません")
            Return False
        End If
        Return True
    End Function


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If IsListboxSelectedValid() = True Then
            ListBox_OrderSet.Items.Insert(ListBox_OrderSet.SelectedIndex, "")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If IsListboxSelectedValid() = True Then
            If ListBox_OrderSet.SelectedIndex + 1 > ListBox_OrderSet.Items.Count - 1 Then
                ListBox_OrderSet.Items.Insert(ListBox_OrderSet.SelectedIndex + 1, "")
            Else
                ListBox_OrderSet.Items.Add("")
            End If
        End If
    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click
        ListBox_OrderSet.SelectedIndex = -1
    End Sub

    Private Sub AddOrderPack(orderstring As String)
        If ListBox_OrderSet.SelectedIndex = -1 Then
            ListBox_OrderSet.Items.Add(orderstring)
            ListBox_OrderSet.ClearSelected()
            MoveCursorToLastpos()
        Else
            'カーソルの下に追加する
            If ListBox_OrderSet.SelectedIndex = ListBox_OrderSet.Items.Count - 1 Then
                ListBox_OrderSet.Items.Add(orderstring)
                ListBox_OrderSet.ClearSelected()
                MoveCursorToLastpos()
            Else
                Dim tmpcursor As Integer = ListBox_OrderSet.SelectedIndex
                ListBox_OrderSet.ClearSelected()
                ListBox_OrderSet.Items.Insert(tmpcursor + 1, orderstring)
                ListBox_OrderSet.SelectedIndex = tmpcursor + 1
            End If
        End If
    End Sub

    Private Sub CopySelectedItemsToClip()
        Dim txt As String = ""
        For Each str As String In ListBox_OrderSet.SelectedItems
            txt = txt & str & ";"
        Next
        My.Computer.Clipboard.Clear()
        My.Computer.Clipboard.SetText(txt)
    End Sub

    Private Sub DeleteSelectedItems()
        Dim i As Integer
        Dim dellist As List(Of Integer) = New List(Of Integer)
        For Each index As Integer In ListBox_OrderSet.SelectedIndices
            dellist.Add(index)
        Next
        dellist.Reverse()
        For Each index As Integer In dellist
            ListBox_OrderSet.Items.RemoveAt(index)
        Next
    End Sub

    Private Sub ListBox_OrderSet_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox_OrderSet.KeyDown
        'コピーの挙動
        If e.KeyCode = Keys.C Then
            If (e.Modifiers And Keys.Control) = Keys.Control Then
                CopySelectedItemsToClip()
            End If
        End If
        'ペーストの挙動・内容チェックしてないので入れこむこと
        If e.KeyCode = Keys.V Then
            If (e.Modifiers And Keys.Control) = Keys.Control Then
                If ListBox_OrderSet.SelectedIndex = -1 Then
                    MessageBox.Show("リストボックス内にペーストを行うための有効なカーソルがありません")
                    Return
                End If
                If ListBox_OrderSet.SelectedIndices.Count > 1 Then
                    MessageBox.Show("単一カーソルの直下にペーストするため、選択を１つにしてください")
                    Return
                End If
                Dim lines As String() = My.Computer.Clipboard.GetText().Split(";")
                Dim sta As Stack(Of String) = New Stack(Of String)
                For Each str As String In lines
                    If str.Length > 0 Then
                        sta.Push(str)
                    End If
                Next
                For i = 0 To sta.Count - 1
                    ListBox_OrderSet.Items.Insert(ListBox_OrderSet.SelectedIndex + 1, sta.Pop())
                Next
            End If
        End If
        'deleteキーの挙動、選択行を消す
        If e.KeyCode = Keys.Delete Then
            DeleteSelectedItems()
        End If
        'ctrl+xの挙動
        If e.KeyCode = Keys.X Then
            If (e.Modifiers And Keys.Control) = Keys.Control Then
                CopySelectedItemsToClip()
                DeleteSelectedItems()
            End If
        End If
    End Sub

    Private Sub Button_SetOrderPackWithoutOperand_Click(sender As Object, e As EventArgs) Handles Button_SetOrderPackWithoutOperand.Click
        If ComboBox1.SelectedItem IsNot Nothing Then
            AddOrderPack(ComboBox1.SelectedItem)
        End If
    End Sub


    Private Function RenameFileEx(oldfilename As String) As Boolean
        Dim index As Integer
        Dim newfilename As String
        For index = 0 To 9999
            newfilename = oldfilename & "." & index.ToString
            If IO.File.Exists(newfilename) = False Then
                Rename(oldfilename, newfilename)
                Return True
            End If
        Next
        Return False
    End Function

    Private Function GetLowLevelOrderString(str As String) As String
        Dim strarray() As String = str.Split(",")
        Dim i As Integer = 0
        Dim tmpi As Integer
        Dim result As String = ""
        For Each tmpstr As String In strarray
            If i = 0 Then
                mDictOrderDescriptionAndID.TryGetValue(tmpstr, tmpi)
                result = tmpi.ToString
            Else
                result = result & "," & tmpstr
            End If
            i = i + 1
        Next
        Return result
    End Function


    'ラベル設定命令のみはチェックしてない。
    Private Sub Button_OutputOrderFile_Click(sender As Object, e As EventArgs) Handles Button_OutputOrderFile.Click
        '        Dim sr As IO.StreamWriter = New IO.StreamWriter()
        If ListBox_OrderSet.Items.Count = 0 Then
            MessageBox.Show("ListBoxの設定項目が0です。ファイル出力には１つ以上の項目設定が必要です")
            Return
        End If
        Dim sfd As SaveFileDialog = New SaveFileDialog
        If sfd.ShowDialog() = DialogResult.OK Then
            If System.IO.File.Exists(sfd.FileName) = True Then
                Dim result As DialogResult = MsgBox("リネームして続行しますか？リネームは連番が使われます",
                              MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Exclamation,
                              "すでに存在するファイル")
                If result = DialogResult.No Or result = DialogResult.Cancel Then
                    Return
                End If
                If RenameFileEx(sfd.FileName) = True Then
                    'ここまで抜けたら正常、何もしないで主流へと。
                Else
                    MessageBox.Show("旧ファイルのリネームに失敗しました")
                    Return
                End If
            Else
                '正常な流れ、何もしないで主流へ
            End If
            Dim parentpath As IO.DirectoryInfo = System.IO.Directory.GetParent(sfd.FileName)
            Dim newfilefullpath As String = parentpath.ToString & "\" & IO.Path.GetFileName(sfd.FileName)
            Dim newlabelfilefullpath As String = parentpath.ToString & "\l_" & IO.Path.GetFileName(sfd.FileName)
            If IO.Path.HasExtension(newfilefullpath) = False Then
                newfilefullpath = newfilefullpath & ".csv"
                '                newlabelfilefullpath = newlabelfilefullpath & ".csv"
            End If
            '主流
            Dim swOrders As IO.StreamWriter = New IO.StreamWriter(newfilefullpath, False, System.Text.Encoding.UTF8)
            For Each tmpstr As String In ListBox_OrderSet.Items
                Dim llstr As String = GetLowLevelOrderString(tmpstr)
                swOrders.WriteLine(llstr)
                'If IsLabelOrder(llstr) Then
                'ラベルのオーダーは別ファイル出力
                'swLabels.WriteLine(llstr)
                'Else
                'End If
            Next
            swOrders.Close()
        End If
    End Sub

    Private Sub Button_LoadOrderFile_Click(sender As Object, e As EventArgs) Handles Button_LoadOrderFile.Click
        Dim ofd As OpenFileDialog = New OpenFileDialog()
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim sr As IO.StreamReader = New IO.StreamReader(ofd.FileName)
            Dim tmpstr As String
            Dim strarray() As String
            Dim desc As String
            Dim outstr As String
            Do While sr.EndOfStream = False
                tmpstr = sr.ReadLine
                strarray = tmpstr.Split(",")
                If mDictOrderIDAndDescription.TryGetValue(Integer.Parse(strarray.GetValue(0)), desc) = True Then
                    outstr = desc
                End If
                Dim i As Integer = 0
                For Each tmp As String In strarray
                    If i = 0 Then
                    Else
                        outstr = outstr & "," & tmp
                    End If
                    i = i + 1
                Next
                ListBox_OrderSet.Items.Add(outstr)
            Loop
            sr.Close()
        End If
    End Sub
End Class
