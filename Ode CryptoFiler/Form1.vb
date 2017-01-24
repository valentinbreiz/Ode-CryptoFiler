Imports System.Net
Imports System.ComponentModel
Imports System.IO
Imports System.Net.Mail
Imports System.Text
Imports System.Security.Cryptography



Public Class Form1
#Region "Variables globales"
    Public VersionActuelle = "0.1 - Beta Test"
    Dim buffer(4096) As Byte
    Dim nomfichier_src As String
    Dim nomfichier_src_ext As String
    Dim fichierstreamRead As FileStream
    Dim fichierstreamWrite As FileStream
    Dim nomfichier_dest As String
    Dim lien As String
    Dim extension As String
    Dim repertoire As String
    Dim cle As Integer

    Dim nomfichier As String

    Dim strFileToEncrypt As String
    Dim strFileToDecrypt As String
    Dim strOutputEncrypt As String
    Dim strOutputDecrypt As String
    Dim fsInput As System.IO.FileStream
    Dim fsOutput As System.IO.FileStream

#End Region



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        txtDestinationDecrypt.Hide()
        txtDestinationEncrypt.Hide()
        ProgressBar1.Hide()
        pnl_retour.Hide()
        pnl_infos.Hide()
        pnl_principal.Show()
        key.Hide()
        Label17.Hide()
        RetourÀLacceuilToolStripMenuItem.Visible = False
        Label9.Text = "Ode CryptoFiler " & VersionActuelle
        EarnButton2.Hide()
        EarnButton1.Hide()
        Me.Text = "Ode CryptoFiler " & VersionActuelle

    End Sub

#Region "Page Infos"
    Private Sub Majsauvegarde_FileOk(sender As Object, e As CancelEventArgs) Handles Majsauvegarde.FileOk
        Dim majpath = Majsauvegarde.FileName
        Dim telechargement As New WebClient
        Dim lientelechargement As String = telechargement.DownloadString("http://partagedefiches.xyz/ode_cryptofiler/telechargement.txt")
        My.Computer.Network.DownloadFile(lientelechargement, majpath)
        MsgBox("Mise à jour effectuée !", MsgBoxStyle.Information, "")
        End
    End Sub

    Private Sub EarnButton4_Click(sender As Object, e As EventArgs) Handles EarnButton4.Click
        pnl_retour.Show()
        pnl_principal.Hide()
        pnl_infos.Hide()
    End Sub

    Private Sub EarnButton5_Click(sender As Object, e As EventArgs) Handles EarnButton5.Click

        Label3.Hide()
        If TextBox3.Text = "" Then
            MsgBox("Vous n'avez pas renseigné de Sujet ! ", MsgBoxStyle.Critical, "Erreur")


        ElseIf TextBox4.Text = "" Then
            MsgBox("Vous n'avez pas renseigné de commentaire ! ", MsgBoxStyle.Critical, "Erreur")
        Else

            Try
                Using client = New WebClient()
                    Using stream = client.OpenRead("http://www.google.com")
                        CheckUpdates()
                        Dim informations As New WebClient
                        Dim pass As New WebClient
                        Dim informationscontenu As String = informations.DownloadString("http://partagedefiches.xyz/ode_cryptofiler/informations.txt")
                        Dim mdp As String = pass.DownloadString("http://partagedefiches.xyz/ode_cryptofiler/pass.php")
                        Label10.Text = informationscontenu

                        Dim mesBytes() As Byte = Convert.FromBase64String(mdp)
                        Dim motdepassedecrypte = Encoding.UTF8.GetString(mesBytes)

                        Dim mail As New MailMessage
                        Dim SMTP As New SmtpClient("smtp.gmail.com")
                        mail.From = New MailAddress("valentinbreiz@gmail.com")
                        mail.To.Add("valentinbreiz@gmail.com")
                        mail.Subject = TextBox3.Text
                        mail.Body = TextBox1.Text & " a envoyé : " & TextBox4.Text
                        SMTP.Port = "587"
                        SMTP.Credentials = New System.Net.NetworkCredential("valentinbreiz@gmail.com", motdepassedecrypte)
                        SMTP.EnableSsl = True

                        SMTP.Send(mail)
                        MsgBox("Retour envoyé !", MsgBoxStyle.Information, "")
                        pnl_retour.Hide()
                        pnl_infos.Show()
                        TextBox3.Clear()
                        TextBox1.Clear()
                        TextBox4.Clear()

                    End Using
                End Using
            Catch
                MsgBox("Pas de connexion Internet, impossible d'envoyer de retour.", MsgBoxStyle.Critical, "Erreur : Pas de connexion Internet")
                Label8.Show()
            End Try
        End If
    End Sub
#Region "Infos_Mario"
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Imageversdroite.Tick
        PictureBox1.Left = PictureBox1.Left + 1
        If PictureBox1.Location.X = 673 Then
            PictureBox1.Hide()
            PictureBox2.Show()
            PictureBox1.Location = New Point(12, 187)
            Imageversdroite.Stop()
            Imageversgauche.Start()
        End If
    End Sub
    Private Sub Imageversgauche_Tick(sender As Object, e As EventArgs) Handles Imageversgauche.Tick
        PictureBox2.Left = PictureBox2.Left - 1
        If PictureBox2.Location.X = 12 Then
            PictureBox2.Hide()
            PictureBox1.Show()
            PictureBox2.Location = New Point(673, 187)
            Imageversgauche.Stop()
            Imageversdroite.Start()
        End If
    End Sub
    Private Sub PictureBox1_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox1.DoubleClick
        MsgBox("Tu as réussi à capturer Mario !", MsgBoxStyle.Information, "Bravo !")
        PictureBox1.Hide()
        PictureBox2.Hide()
    End Sub

    Private Sub PictureBox2_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox2.DoubleClick
        MsgBox("Tu as réussi à capturer Mario !", MsgBoxStyle.Information, "Bravo !")
        PictureBox1.Hide()
        PictureBox2.Hide()
    End Sub
#End Region
    Private Sub verifier_Click(sender As Object, e As EventArgs) Handles verifier.Click
        Label3.Hide()
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://www.google.com")
                    CheckUpdates()
                    Dim informations As New WebClient
                    Dim informationscontenu As String = informations.DownloadString("http://partagedefiches.xyz/ode_cryptofiler/informations.txt")
                    Label10.Text = informationscontenu
                End Using
            End Using
        Catch
            MsgBox("Pas de connexion Internet, impossible de trouver des mises à jours.", MsgBoxStyle.Critical, "Erreur : Pas de connexion Internet")
            Label8.Show()
        End Try
    End Sub
    Sub CheckUpdates()
        Dim MAJ As New WebClient
        Dim DerniereVersion As String = MAJ.DownloadString("http://partagedefiches.xyz/ode_cryptofiler/version.txt")
        If VersionActuelle = DerniereVersion Then
            Label3.Hide()
            Label4.Show()
            verifier.Focus()

        ElseIf VersionActuelle > DerniereVersion Then
            Label13.Show()

        Else
            Label3.Hide()
            verifier.Hide()
            Label5.Show()
            EarnButton3.Show()
            EarnButton3.Focus()


        End If
    End Sub
    Private Sub EarnButton3_Click(sender As Object, e As EventArgs) Handles EarnButton3.Click
        Dim version As New WebClient
        Dim nomversion As String = version.DownloadString("http://partagedefiches.xyz/ode_cryptofiler/version.txt")
        Majsauvegarde.FileName = "Ode CryptoFiler - " & nomversion
        Majsauvegarde.ShowDialog()
    End Sub

    Private Sub RetourÀLacceuilToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RetourÀLacceuilToolStripMenuItem.Click
        pnl_infos.Hide()
        pnl_retour.Hide()
        pnl_principal.Show()
        RetourÀLacceuilToolStripMenuItem.Visible = False
        FichierToolStripMenuItem.Visible = True
        ÀProposToolStripMenuItem.Visible = True
    End Sub
#End Region

#Region "Créer une clé "

    Private Function CreateKey(ByVal strPassword As String) As Byte()

        Dim chrData() As Char = strPassword.ToCharArray

        Dim intLength As Integer = chrData.GetUpperBound(0)

        Dim bytDataToHash(intLength) As Byte

        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next

        Dim SHA512 As New System.Security.Cryptography.SHA512Managed

        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)

        Dim bytKey(31) As Byte

        For i As Integer = 0 To 31
            bytKey(i) = bytResult(i)
        Next

        Return bytKey
    End Function

#End Region

#Region "Créer un IV "

    Private Function CreateIV(ByVal strPassword As String) As Byte()

        Dim chrData() As Char = strPassword.ToCharArray

        Dim intLength As Integer = chrData.GetUpperBound(0)

        Dim bytDataToHash(intLength) As Byte

        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next

        Dim SHA512 As New System.Security.Cryptography.SHA512Managed

        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)

        Dim bytIV(15) As Byte


        For i As Integer = 32 To 47
            bytIV(i - 32) = bytResult(i)
        Next

        Return bytIV
    End Function

#End Region

#Region "Crypter décrypter un fichier "
    Private Enum CryptoAction
        ActionEncrypt = 1
        ActionDecrypt = 2
    End Enum

    Private Sub EncryptOrDecryptFile(ByVal strInputFile As String,
                                     ByVal strOutputFile As String,
                                     ByVal bytKey() As Byte,
                                     ByVal bytIV() As Byte,
                                     ByVal Direction As CryptoAction)
        Try


            fsInput = New System.IO.FileStream(strInputFile, FileMode.Open,
                                               FileAccess.Read)
            fsOutput = New System.IO.FileStream(strOutputFile, FileMode.OpenOrCreate,
                                                FileAccess.Write)
            fsOutput.SetLength(0)

            Dim bytBuffer(4096) As Byte
            Dim lngBytesProcessed As Long = 0
            Dim lngFileLength As Long = fsInput.Length
            Dim intBytesInCurrentBlock As Integer
            Dim csCryptoStream As CryptoStream

            Dim cspRijndael As New System.Security.Cryptography.RijndaelManaged

            ProgressBar1.Value = 0
            ProgressBar1.Maximum = 100


            Select Case Direction
                Case CryptoAction.ActionEncrypt
                    csCryptoStream = New CryptoStream(fsOutput,
                    cspRijndael.CreateEncryptor(bytKey, bytIV),
                    CryptoStreamMode.Write)

                Case CryptoAction.ActionDecrypt
                    csCryptoStream = New CryptoStream(fsOutput,
                    cspRijndael.CreateDecryptor(bytKey, bytIV),
                    CryptoStreamMode.Write)
            End Select

            While lngBytesProcessed < lngFileLength

                intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, 4096)

                csCryptoStream.Write(bytBuffer, 0, intBytesInCurrentBlock)

                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)

                ProgressBar1.Value = CInt((lngBytesProcessed / lngFileLength) * 100)
            End While

            csCryptoStream.Close()
            fsInput.Close()
            fsOutput.Close()

            If Direction = CryptoAction.ActionEncrypt Then
                Dim fileOriginal As New FileInfo(strFileToEncrypt)
                fileOriginal.Delete()
            End If

            If Direction = CryptoAction.ActionDecrypt Then
                Dim fileEncrypted As New FileInfo(strFileToDecrypt)
                fileEncrypted.Delete()
            End If

            Dim Wrap As String = Chr(13) + Chr(10)
            Dim lngKBytesProcessed = lngBytesProcessed / 1000
            If Direction = CryptoAction.ActionEncrypt Then
                MsgBox("Cryptage terminé !" + Wrap + Wrap +
                        "Nombre total de kilo-octets : " +
                        lngKBytesProcessed.ToString,
                        MsgBoxStyle.Information, "Ode CryptoFiler")

                ProgressBar1.Value = 0
                key.Text = ""
                txtDestinationEncrypt.Text = ""
                EarnButton2.Hide()
                ProgressBar1.Hide()
                key.Text = "Rn5T6wMf935Zfe5FYfd4"

            Else

                MsgBox("Décryptage terminé !" + Wrap + Wrap +
                       "Nombre total de kilo-octets : " +
                        lngKBytesProcessed.ToString,
                        MsgBoxStyle.Information, "Ode CryptoFiler")

                ProgressBar1.Value = 0
                key.Text = ""
                txtDestinationDecrypt.Text = ""
                EarnButton1.Hide()
                ProgressBar1.Hide()
                key.Text = "Rn5T6wMf935Zfe5FYfd4"
            End If


        Catch When Err.Number = 53
            MsgBox("Assurez-vous que le chemin d'accès et le nom de fichier" +
                    "sont correctes et que le fichier existe.",
                     MsgBoxStyle.Exclamation, "Chemin d'accès ou nom de fichier non valide")

        Catch
            fsInput.Close()
            fsOutput.Close()

            If Direction = CryptoAction.ActionDecrypt Then
                Dim fileDelete As New FileInfo(txtDestinationDecrypt.Text)
                fileDelete.Delete()
                ProgressBar1.Value = 0

                MsgBox("Êtes vous sûr d'avoir renseigné la bonne clé de cryptage ?", MsgBoxStyle.Exclamation, "Clé invalide")
            Else
                Dim fileDelete As New FileInfo(txtDestinationEncrypt.Text)
                fileDelete.Delete()

                ProgressBar1.Value = 0

                MsgBox("Ce fichier ne peut pas être crypté.", MsgBoxStyle.Exclamation, "Fichier invalide")

            End If

        End Try
    End Sub

#End Region



    Private Sub ÀProposToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ÀProposToolStripMenuItem.Click
        Label13.Hide()
        Label3.Show()
        verifier.Show()
        pnl_infos.Show()
        pnl_principal.Hide()
        RetourÀLacceuilToolStripMenuItem.Visible = True
        FichierToolStripMenuItem.Visible = False
        ÀProposToolStripMenuItem.Visible = False
        Imageversdroite.Start()
        PictureBox2.Hide()
        Label8.Hide()
        Label4.Hide()
        Label5.Hide()
        EarnButton3.Hide()
        verifier.Focus()
    End Sub

    Private Sub QuitterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitterToolStripMenuItem.Click
        End
    End Sub

    Private Sub ModifierCléDeCryptageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ModifierCléDeCryptageToolStripMenuItem.Click
        key.Show()
        Label17.Show()
    End Sub

    Private Sub ProtegerUnFichierToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProtegerUnFichierToolStripMenuItem.Click
        key.Hide()
        Label17.Hide()

        OpenFileDialog.FileName = ""
        OpenFileDialog.Title = "Choix du fichier à crypter"
        OpenFileDialog.Filter = "Tout type de fichier (*.*) | *.*"

        If OpenFileDialog.ShowDialog = DialogResult.OK Then
            strFileToEncrypt = OpenFileDialog.FileName
            nomfichier = OpenFileDialog.SafeFileName

            Dim add As New ListViewItem(nomfichier)
            add.SubItems.Add(strFileToEncrypt)
            add.SubItems.Add(DateTime.Now.Date & " à " & DateTime.Now.ToShortTimeString())
            ListView1.Items.Add(add)

            Dim iPosition As Integer = 0
            Dim i As Integer = 0

            While strFileToEncrypt.IndexOf("\"c, i) <> -1
                iPosition = strFileToEncrypt.IndexOf("\"c, i)
                i = iPosition + 1
            End While


            strOutputEncrypt = strFileToEncrypt.Substring(iPosition + 1)

            Dim S As String = strFileToEncrypt.Substring(0, iPosition + 1)

            txtDestinationEncrypt.Text = S + strOutputEncrypt + ".ode"

            EarnButton2.Show()
            EarnButton2.Focus()
        End If
    End Sub

    Private Sub OuvrirFichierSécuriserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OuvrirFichierSécuriserToolStripMenuItem.Click

        OpenFileDialog.FileName = ""
        OpenFileDialog.Title = "Choix du fichier à decrypter"
        OpenFileDialog.Filter = "Fichier crypté (*.ode) | *.ode"

        If OpenFileDialog.ShowDialog = DialogResult.OK Then
            strFileToDecrypt = OpenFileDialog.FileName
            nomfichier = OpenFileDialog.SafeFileName

            Dim add As New ListViewItem(nomfichier)
            add.SubItems.Add(strFileToDecrypt)
            add.SubItems.Add(DateTime.Now.Date & " à " & DateTime.Now.ToShortTimeString())
            ListView1.Items.Add(add)
            Dim iPosition As Integer = 0
            Dim i As Integer = 0


            While strFileToDecrypt.IndexOf("\"c, i) <> -1
                iPosition = strFileToDecrypt.IndexOf("\"c, i)
                i = iPosition + 1
            End While


            strOutputDecrypt = strFileToDecrypt.Substring(0, strFileToDecrypt.Length - 4)

            Dim S As String = strFileToDecrypt.Substring(0, iPosition + 1)

            strOutputDecrypt = strOutputDecrypt.Substring((iPosition + 1))

            txtDestinationDecrypt.Text = S + strOutputDecrypt.Replace("_"c, "."c)

            EarnButton1.Show()
            EarnButton1.Focus()

        End If
    End Sub

    Private Sub EarnButton1_Click(sender As Object, e As EventArgs) Handles EarnButton1.Click
        ProgressBar1.Show()

        Dim bytKey As Byte()
        Dim bytIV As Byte()

        bytKey = CreateKey(key.Text)

        bytIV = CreateIV(key.Text)

        EncryptOrDecryptFile(strFileToDecrypt, txtDestinationDecrypt.Text,
                             bytKey, bytIV, CryptoAction.ActionDecrypt)
    End Sub

    Private Sub EarnButton2_Click(sender As Object, e As EventArgs) Handles EarnButton2.Click


        ProgressBar1.Show()


        Dim bytKey As Byte()
        Dim bytIV As Byte()

        bytKey = CreateKey(key.Text)

        bytIV = CreateIV(key.Text)

        EncryptOrDecryptFile(strFileToEncrypt, txtDestinationEncrypt.Text,
                                 bytKey, bytIV, CryptoAction.ActionEncrypt)
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("https://github.com/valentinbreiz")
    End Sub
End Class
