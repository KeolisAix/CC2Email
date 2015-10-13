Imports System.Net.Mail
Imports System.Text

Public Class Form1
    Dim FichierCCVALIDATOR As Boolean = True
    Dim Donnees = "C:\COMPTAGE\DONNEES\" 'Variable de des données
    Dim Evenement = "C:\COMPTAGE\EVENEMENTS\DOSSIER_ALARME_PATRICE\VIDAGE_BUS_2.csv" 'Variable  du Vidage
    Dim IlExisteDepot1 As String
    Dim IlExisteDepot2 As String
    Dim pb As String
    Dim pb2 As String
    Dim TotalNoDonnees As Integer = 0
    Dim TotalNoVidage As Integer = 0
    Dim LignePlus10 As String = ""
    Dim LignePasInfo As String = ""
    Dim DateDuJour As DateTime = Date.Now
    Dim format As String = "yyMMdd"
    Dim sb As New StringBuilder

    Dim FichierDuJour As String = "1_" + DateDuJour.ToString(Format) + ".csv"
    Dim FichierDuJour2 As String = "2_" + DateDuJour.ToString(Format) + ".csv"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If FichierCCVALIDATOR = False Then
            Dim i As Integer
            For i = 0 To My.Computer.FileSystem.GetFiles(Donnees).Count - 1
                TextBox3.AppendText((My.Computer.FileSystem.GetFiles(Donnees).Item(i)) & Environment.NewLine)
            Next i
            DateDuJour = DateDuJour.AddDays(-1)
            If My.Computer.FileSystem.FileExists(Donnees + FichierDuJour.ToString) Then
                IlExisteDepot1 = "Oui"
            Else
                IlExisteDepot1 = "Non"
            End If
            If My.Computer.FileSystem.FileExists(Donnees + FichierDuJour2.ToString) Then
                IlExisteDepot2 = "Oui"
            Else
                IlExisteDepot2 = "Non"
            End If
            If My.Computer.FileSystem.FileExists(Evenement) Then My.Computer.FileSystem.DeleteFile(Evenement)
            Try
                My.Computer.FileSystem.CopyFile("C:\COMPTAGE\EVENEMENTS\VIDAGE_BUS.csv", Evenement)
            Catch supp As Exception
                MsgBox("Impossible à copier : " + supp.ToString)
            End Try
            Try
                pb = "reset"
                For Each line As String In System.IO.File.ReadAllLines(Evenement)
                    Dim word = line.Split(";")
                    DataGridView1.Rows.Add(word(0), word(1), word(2), word(3), word(4))

                    If Val(word(2)) >= "5" And word(2) <> "Nb jours sans vidage" Then
                        pb = "Oui"
                        LignePlus10 = LignePlus10 + "Le BUS : <b>" + word(0) + "</b> n'a pas vidé depuis le <b>" + word(1) + "</b> soit <b>" + word(2) + "</b> jours.<br>"
                        TotalNoVidage = TotalNoVidage + 1
                    End If
                    If Val(word(2)) = Nothing And word(2) <> "Nb jours sans vidage" And word(2) <> "0" Then
                        pb2 = "Oui"
                        LignePasInfo = LignePasInfo + "Le BUS : <b>" + word(0) + "</b> n'a pas de données ! <br>"
                        TotalNoDonnees = TotalNoDonnees + 1
                    End If
                Next
            Catch ex As Exception
                pb = "<h2>Probleme avec le fichier Vidage</h2>"
            End Try
            If pb <> "Oui" Then pb = "Non"
        Else
            Evenement = "C:\COMPTAGE\EVENEMENTS\CONTROLE\VIDAGE_CC_FINAL.csv"
            Try
                pb = "Non"
                pb2 = "Non"
                For Each line As String In System.IO.File.ReadAllLines(Evenement)
                    Dim word = line.Split(";")
                    DataGridView1.Rows.Add(word(0), word(1), word(2))
                    Dim now = Date.Today
                    If word(0) = "BUS" Then
                    Else
                        Dim compare = ""
                        If word(1) = Nothing Then
                        Else
                            compare = DateDiff(DateInterval.Day, Date.Parse(word(1)), now, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFullWeek)
                        End If
                        If Val(compare) >= "5" And word(2) <> "Nb jours sans vidage" Then
                            pb = "Oui"
                            LignePlus10 = LignePlus10 + "Le BUS : <b>" + word(0) + "</b> n'a pas vidé depuis le <b>" + word(1) + "</b> soit <b>" + compare.ToString + "</b> jours.<br>"
                            TotalNoVidage = TotalNoVidage + 1
                        End If
                        If Val(compare) = Nothing And compare <> "Nb jours sans vidage" And compare <> "0" Then
                            pb2 = "Oui"
                            LignePasInfo = LignePasInfo + "Le BUS : <b>" + word(0) + "</b> n'a pas de données ! <br>"
                            TotalNoDonnees = TotalNoDonnees + 1
                        End If
                    End If
                Next
            Catch ex As Exception
                pb = "<h2>Probleme avec le fichier Vidage</h2>"
            End Try
        End If
        Try
            If FichierCCVALIDATOR = False Then
                sb.AppendLine("<center><img src='K:\Commun\Keolis_logo.png'><br><br>Date du rapport : <b>" + Date.Today + "</b> <br>Existance du Fichier du jour pour le DEPOT 1 : <b>" + IlExisteDepot1.ToString + "</b> <br>Existance du Fichier du jour pour le DEPOT 2 : <b>" + IlExisteDepot2.ToString + "</b> <br>Plus de 5 jours sans vidage (Total : " + TotalNoVidage.ToString + ") : <b>" + pb.ToString + "</b> <br>Absence de données (Total : " + TotalNoDonnees.ToString + ") : <b>" + pb2.ToString + "</b><br></center>")
            Else
                sb.AppendLine("<center><img src='K:\Commun\Keolis_logo.png'><br><br>Date du rapport : <b>" + Date.Today + "</b><br>Plus de 5 jours sans vidage (Total : " + TotalNoVidage.ToString + ") : <b>" + pb.ToString + "</b> <br>Absence de données (Total :  " + TotalNoDonnees.ToString + " ) : <b>" + pb2.ToString + "</b><br></center>")
            End If
        Catch aa As Exception
            MsgBox("pb")
        End Try
        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.EnableSsl = True
            Smtp_Server.UseDefaultCredentials = True
            Smtp_Server.Credentials = New Net.NetworkCredential("cellules_compt", "123456789")
            Smtp_Server.DeliveryMethod = SmtpDeliveryMethod.Network
            Smtp_Server.Port = 25
            Smtp_Server.Host = "webmail.keolis.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("cellules_compt@keolis.com")
            e_mail.To.Add("patrice.maldi@keolis.com")
            e_mail.CC.Add("sylvain.mennillo@keolis.com")
            e_mail.CC.Add("patrick.schinle@keolis.com")
            e_mail.Subject = "Recap Cellules Compteuses du " + Date.Now + ""
            e_mail.IsBodyHtml = True
            If FichierCCVALIDATOR = False Then
                If IlExisteDepot1 = "Oui" Then e_mail.Attachments.Add(New Attachment(Donnees + FichierDuJour.ToString))
                If IlExisteDepot2 = "Oui" Then e_mail.Attachments.Add(New Attachment(Donnees + FichierDuJour2.ToString))
            End If
            If pb = "Oui" Or pb2 = "Oui" Then
                e_mail.Attachments.Add(New Attachment(Evenement))
                sb.AppendLine("<br><br><center>" + LignePlus10 + "</center><br>")
                sb.AppendLine("<br><br><center>" + LignePasInfo + "</center>")
            End If
            e_mail.Body = sb.ToString()
            Smtp_Server.Send(e_mail)
            Me.Dispose()

        Catch error_t As Exception
            MsgBox("Une Erreur lors de l'envoi de Mail :" + error_t.ToString)
        End Try
    End Sub
End Class
