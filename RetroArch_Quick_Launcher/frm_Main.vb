Imports System.ComponentModel
Imports System.Text
Imports TC = MKNetLib.cls_MKTypeConverter

Partial Public Class frm_Main
	Inherits MKNetDXLib.frm_MKDXBaseForm

	Public dataPath As String = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%") & "\.retroarch-quick-launcher"
	Public configPath As String = dataPath & "\" & "config.xml"

    Private launchImmediately As Boolean = False

    Shared Sub New()
        DevExpress.UserSkins.BonusSkins.Register()
        DevExpress.Skins.SkinManager.EnableFormSkins()
    End Sub

    Public Sub New()
        InitializeComponent()

        MKNetDXLib.frm_MKDXBaseForm.Default_Form_Icon = Me.Icon
        MKNetDXLib.cls_MKDXSkin.LoadSkin("DevExpress Dark Style")

        If Not Alphaleonis.Win32.Filesystem.Directory.Exists(Me.dataPath) Then
            Alphaleonis.Win32.Filesystem.Directory.CreateDirectory(Me.dataPath)
        End If

        If Not Alphaleonis.Win32.Filesystem.Directory.Exists(Me.dataPath) Then
            MKNetDXLib.cls_MKDXHelper.MessageBox("Unable to create directory '" & dataPath & "'!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Me.dataPath = ""
            Me.configPath = ""
        End If

        If Alphaleonis.Win32.Filesystem.File.Exists(configPath) Then
            Me.DS.ReadXml(configPath)
        Else
            Me.DS.tbl_CurrentConfig.Rows.Add(Me.DS.tbl_CurrentConfig.NewRow)
        End If

        Me.BS_CurrentConfig.EndEdit()

        Dim cmdArgs = Environment.GetCommandLineArgs()
        For i As Integer = 1 To cmdArgs.Length - 1
            Me.ImportFile(cmdArgs(i), enm_ImportFileMode.AUTODETECT)
        Next
    End Sub

    Public Enum enm_ImportFileMode
		GAME = 1
		PATCH = 2
		AUTODETECT = 3
	End Enum

    Public Sub ImportFile(ByVal path As String, ByVal mode As enm_ImportFileMode)
        Dim target As DevExpress.XtraEditors.TextEdit = Nothing

        If mode = enm_ImportFileMode.PATCH OrElse (mode = enm_ImportFileMode.AUTODETECT AndAlso {".ips", ".ups", ".bps"}.Contains(Alphaleonis.Win32.Filesystem.Path.GetExtension(path).ToLower)) Then
            Me.txb_Patch_Location.EditValue = path
        Else
            'TODO: .zip file handling
            Me.txb_Game_Location.EditValue = path

            Load_FileExtensionConfig()
        End If

    End Sub

    Private Sub btn_Browse_Game_Click(sender As Object, e As EventArgs) Handles btn_Browse_Game.Click
		Dim ofd = New System.Windows.Forms.OpenFileDialog()
		Dim dialogResult = ofd.ShowDialog()
		If dialogResult <> DialogResult.OK Then
			Return
		End If

		Me.ImportFile(ofd.FileName, enm_ImportFileMode.GAME)
	End Sub

	Private Sub btn_Browse_Patch_Click(sender As Object, e As EventArgs) Handles btn_Browse_Patch.Click
		Dim ofd = New System.Windows.Forms.OpenFileDialog()
		ofd.Filter = "Supported Patch Files (*.ips;*.ups;*.bps)|*.ips;*.ups;*.bps"
		Dim dialogResult = ofd.ShowDialog()
		If dialogResult <> DialogResult.OK Then
			Return
		End If

		Me.ImportFile(ofd.FileName, enm_ImportFileMode.PATCH)
	End Sub

	Private Sub setRetroArchLocation(path)
		Me.BS_CurrentConfig.Current("RetroArch_Location") = path
		Me.BS_CurrentConfig.EndEdit()

		Refill_Libretro_Cores()
		Refill_Shader()
	End Sub

	Private Sub btn_Browse_RetroArch_Click(sender As Object, e As EventArgs) Handles btn_Browse_RetroArch.Click
		Dim ofd = New System.Windows.Forms.OpenFileDialog()
		ofd.Filter = "RetroArch Executable (retroarch*.exe)|retroarch*.exe"
		Dim dialogResult = ofd.ShowDialog()
		If dialogResult <> DialogResult.OK Then
			Return
		End If

		setRetroArchLocation(ofd.FileName)
	End Sub

	Private Sub Refill_Libretro_Cores()
		Dim oLastCore As Object = BS_CurrentConfig.Current("Libretro_Core")
		Me.DS.tbl_Libretro_Cores.Clear()

		If Not Alphaleonis.Win32.Filesystem.File.Exists(Me.txb_RetroArch_Location.Text) Then
			Return
		End If

		Cursor.Current = Cursors.WaitCursor

		Dim basedir As String = MKNetLib.cls_MKStringSupport.Clean_Right(Alphaleonis.Win32.Filesystem.Path.GetDirectoryName(Me.txb_RetroArch_Location.Text), "\")
		Dim coredir As String = basedir & "\cores"
		Dim infodir As String = basedir & "\info"

		If Alphaleonis.Win32.Filesystem.Directory.Exists(coredir) Then
			Dim files As String() = Alphaleonis.Win32.Filesystem.Directory.GetFiles(coredir, "*", IO.SearchOption.TopDirectoryOnly)
			For Each file As String In files
				If Alphaleonis.Win32.Filesystem.Path.GetExtension(file).ToLower = ".dll" Then
					Dim row_file As DataRow = Me.DS.tbl_Libretro_Cores.NewRow
					row_file("DLL") = Alphaleonis.Win32.Filesystem.Path.GetFileName(file)
					row_file("Displayname") = Alphaleonis.Win32.Filesystem.Path.GetFileName(file)

					'TODO: find Displayname in RetroArch DB if possible
					Dim infofile As String = infodir & "\" & Alphaleonis.Win32.Filesystem.Path.GetFileNameWithoutExtension(file) & ".info"
					If Alphaleonis.Win32.Filesystem.File.Exists(infofile) Then
						Dim sContent As String = MKNetLib.cls_MKFileSupport.GetFileContents(infofile)
						If Not TC.IsNullNothingOrEmpty(sContent) Then
							For Each line As String In sContent.Split(vbLf)
								If line.Contains("display_name") Then
									If MKNetLib.cls_MKRegex.IsMatch(line, """(.*?)""") Then
										row_file("Displayname") = MKNetLib.cls_MKRegex.GetMatches(line, """(.*?)""")(0).Groups(1).Captures(0).Value
									End If
								End If
							Next
						End If
					End If

					Me.DS.tbl_Libretro_Cores.Rows.Add(row_file)
				End If
			Next
		End If

		If Me.BS_Libretro_Cores.Find("DLL", oLastCore) >= 0 Then
			Me.cmb_Libretro_Core.EditValue = oLastCore
			SetBindingSourcePosition(Me.BS_Libretro_Cores, "DLL", oLastCore)
		End If

		Cursor.Current = Cursors.Default
	End Sub

	Public Function SetBindingSourcePosition(ByRef BS As BindingSource, ByVal ColumnName As String, ByVal Value As Object) As Boolean
		Dim properties As System.ComponentModel.PropertyDescriptorCollection
		Dim _property As System.ComponentModel.PropertyDescriptor

		Try
			properties = CType(Me, System.ComponentModel.ITypedList).GetItemProperties(Nothing)
			_property = properties(ColumnName)
			If _property IsNot Nothing Then
				Dim iPos As Integer = BS.Find(_property, Value)
				BS.Position = BS.Find(_property, Value)

				If iPos > -1 Then
					Return True
				Else
					Return False
				End If
			Else
				Return False
			End If
		Catch
			'
		End Try
	End Function

	Private Sub Refill_Shader()
		Dim oLastShader As Object = BS_CurrentConfig.Current("Shader")
		Me.DS.tbl_Shader.Clear()

		If Not Alphaleonis.Win32.Filesystem.File.Exists(Me.txb_RetroArch_Location.Text) Then
			Return
		End If

		Cursor.Current = Cursors.WaitCursor

		Dim basedir As String = MKNetLib.cls_MKStringSupport.Clean_Right(Alphaleonis.Win32.Filesystem.Path.GetDirectoryName(Me.txb_RetroArch_Location.Text), "\")
		Dim shaderdir As String = basedir & "\shaders\"

		If Alphaleonis.Win32.Filesystem.Directory.Exists(shaderdir) Then
			Dim fsrch As New MKNetLib.cls_MKFileSearch(New Alphaleonis.Win32.Filesystem.DirectoryInfo(shaderdir))
			fsrch.Search(New Alphaleonis.Win32.Filesystem.DirectoryInfo(shaderdir), "*.glslp")
			fsrch.Search(New Alphaleonis.Win32.Filesystem.DirectoryInfo(shaderdir), "*.slangp")

			Dim arrFiles As New System.Collections.ArrayList
			arrFiles.AddRange(fsrch.Files)

			For Each fi As Alphaleonis.Win32.Filesystem.FileInfo In arrFiles
				Dim row_file As DataRow = Me.DS.tbl_Shader.NewRow
				row_file("Path") = fi.FullName.Replace(shaderdir, "")
				Me.DS.tbl_Shader.Rows.Add(row_file)
			Next
		End If

		If Me.BS_Shader.Find("Path", oLastShader) >= 0 Then
			cmb_Retroarch_Shader.EditValue = oLastShader
			SetBindingSourcePosition(Me.BS_Shader, "Path", oLastShader)
		End If

		Cursor.Current = Cursors.Default
	End Sub

	Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
		Me.Close()
	End Sub

    Private Sub Load_FileExtensionConfig()
        Dim FileExtension = Alphaleonis.Win32.Filesystem.Path.GetExtension(Me.txb_Game_Location.Text).ToLower.Replace(".", "")

        Dim row As DS_RetroArch_Quick_Launcher.tbl_FileExtensionConfigRow = Nothing
        Dim rows = DS.tbl_FileExtensionConfig.Select("FileExtension = " & TC.getSQLParameter(FileExtension))

        If rows.Length = 0 Then
            Return
        End If

        Me.BS_CurrentConfig.Current("Libretro_Core") = rows(0)("Libretro_Core")
        Me.BS_CurrentConfig.Current("Shader") = rows(0)("Shader")
        Me.chb_LaunchImmediately.Checked = TC.NZ(rows(0)("LaunchImmediately"), False)
    End Sub

    Private Sub save_FileExtensionConfig()
        Dim row As DS_RetroArch_Quick_Launcher.tbl_FileExtensionConfigRow = Nothing
        Dim isNew As Boolean = False

        Dim FileExtension = Alphaleonis.Win32.Filesystem.Path.GetExtension(Me.txb_Game_Location.Text).ToLower.Replace(".", "")

        Dim rows = DS.tbl_FileExtensionConfig.Select("FileExtension = " & TC.getSQLParameter(FileExtension))

        If rows.Length > 0 Then
            row = rows(0)
        Else
            isNew = True
            row = DS.tbl_FileExtensionConfig.NewRow
        End If

        row.FileExtension = FileExtension
        row("Libretro_Core") = Me.cmb_Libretro_Core.EditValue
        row("Shader") = Me.cmb_Retroarch_Shader.EditValue
        row("LaunchImmediately") = Me.chb_LaunchImmediately.Checked

        If isNew Then
            Me.DS.tbl_FileExtensionConfig.Rows.Add(row)
        End If
    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
		If Not Alphaleonis.Win32.Filesystem.File.Exists(Me.txb_RetroArch_Location.EditValue) Then
            MKNetDXLib.cls_MKDXHelper.MessageBox("RetroArch executable not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
		End If

        If Not Alphaleonis.Win32.Filesystem.File.Exists(Me.txb_Game_Location.EditValue) Then
            MKNetDXLib.cls_MKDXHelper.MessageBox("Game not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If TC.NZ(Me.cmb_Libretro_Core.EditValue, "") = "" Then
            MKNetDXLib.cls_MKDXHelper.MessageBox("No Libretro Core selected!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim emufullpath = Me.txb_RetroArch_Location.EditValue
		Dim Args As String = ""

		If Alphaleonis.Win32.Filesystem.File.Exists(txb_Patch_Location.Text) Then
			Dim patchType = ""
			Select Case Alphaleonis.Win32.Filesystem.Path.GetExtension(txb_Patch_Location.Text).ToLower
				Case ".ips"
					patchType = "ips"
				Case ".ups"
					patchType = "ups"
				Case ".bps"
					patchType = "bps"
			End Select

			If Not TC.IsNullNothingOrEmpty(patchType) Then
				Args = "--" & patchType & " """ & txb_Patch_Location.Text & """" & (IIf(Args <> "", " ", "")) & Args
			End If
		End If

		If Not TC.IsNullNothingOrEmpty(BS_CurrentConfig.Current("Shader")) Then
			Dim Retroarch_Shader As String = BS_CurrentConfig.Current("Shader").Trim
			Args = "--set-shader """ & Retroarch_Shader & """" & (IIf(Args <> "", " ", "")) & Args
		End If

		If Not TC.IsNullNothingOrEmpty(BS_CurrentConfig.Current("Libretro_Core")) Then
			Dim libretroCore As String = BS_CurrentConfig.Current("Libretro_Core").Trim
			Args = "-L cores\" & libretroCore & (IIf(Args <> "", " ", "")) & Args
		End If

		Args &= IIf(Args <> "", " ", "") & """" & Me.txb_Game_Location.Text & """"

		Dim proc = New System.Diagnostics.Process
		proc.StartInfo.FileName = emufullpath
		proc.StartInfo.WorkingDirectory = Alphaleonis.Win32.Filesystem.Path.GetDirectoryName(emufullpath)
		proc.StartInfo.Arguments = Args
		proc.StartInfo.UseShellExecute = True
		proc.Start()

        save_FileExtensionConfig()

        If Alphaleonis.Win32.Filesystem.Directory.Exists(Me.dataPath) Then
            Me.DS.WriteXml(configPath)
        End If

        Me.Close()
	End Sub

	Private Sub cmb_Libretro_Core_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles cmb_Libretro_Core.ButtonClick
		If e.Button.Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Delete Then
			cmb_Libretro_Core.EditValue = DBNull.Value
		End If
	End Sub

	Private Sub cmb_Retroarch_Shader_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles cmb_Retroarch_Shader.ButtonClick
		Try
			If e.Button.Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Delete Then
				cmb_Retroarch_Shader.EditValue = DBNull.Value
			End If
		Catch ex As Exception

		End Try
	End Sub

	Private Sub txb_Game_Location_DragEnter(sender As Object, e As DragEventArgs) Handles txb_Game_Location.DragEnter
		Try
			Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

			If Not {".ips", ".ups", ".bps"}.Contains(Alphaleonis.Win32.Filesystem.Path.GetExtension(files(0).ToLower)) Then
				e.Effect = DragDropEffects.All
			Else
				e.Effect = DragDropEffects.None
			End If

		Catch ex As Exception

		End Try
	End Sub

	Private Sub txb_Game_Location_DragDrop(sender As Object, e As DragEventArgs) Handles txb_Game_Location.DragDrop
		Try
			Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

            Me.txb_Game_Location.EditValue = files(0)

            Load_FileExtensionConfig()
        Catch ex As Exception

		End Try
	End Sub

	Private Sub txb_RetroArch_Location_DragEnter(sender As Object, e As DragEventArgs) Handles txb_RetroArch_Location.DragEnter
		Try
			Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

			Dim filename = Alphaleonis.Win32.Filesystem.Path.GetFileName(files(0)).ToLower

			If MKNetLib.cls_MKRegex.IsMatch(filename, "retroarch.*?\.exe") Then
				e.Effect = DragDropEffects.All
			Else
				e.Effect = DragDropEffects.None
			End If
		Catch ex As Exception

		End Try
	End Sub

	Private Sub txb_RetroArch_Location_DragDrop(sender As Object, e As DragEventArgs) Handles txb_RetroArch_Location.DragDrop
		Try
			Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

			setRetroArchLocation(files(0))
		Catch ex As Exception

		End Try
	End Sub

	Private Sub txb_Patch_Location_DragEnter(sender As Object, e As DragEventArgs) Handles txb_Patch_Location.DragEnter
		Try
			Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

			If {".ips", ".ups", ".bps"}.Contains(Alphaleonis.Win32.Filesystem.Path.GetExtension(files(0).ToLower)) Then
				e.Effect = DragDropEffects.All
			Else
				e.Effect = DragDropEffects.None
			End If
		Catch ex As Exception

		End Try
	End Sub

	Private Sub txb_Patch_Location_DragDrop(sender As Object, e As DragEventArgs) Handles txb_Patch_Location.DragDrop
		Try
			Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

			Me.txb_Patch_Location.EditValue = files(0)
		Catch ex As Exception

		End Try
	End Sub

	Private Sub lbl_Metropolis_Launcher_Click(sender As Object, e As EventArgs) Handles lbl_Metropolis_Launcher.Click
		Process.Start("https://metropolis-launcher.net")
	End Sub

    Private Sub frm_Main_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If Me.chb_LaunchImmediately.Checked AndAlso Not My.Computer.Keyboard.ShiftKeyDown Then
            'Me.Visible = False
            Me.btn_OK_Click(Me.btn_OK, New EventArgs)
        End If
    End Sub
End Class
