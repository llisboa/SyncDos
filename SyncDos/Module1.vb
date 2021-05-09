Imports Icraft.IcftBase

Module Module1
    Dim Recurso As String = ""
    Dim EmailPara As String = ""
    Dim Servidor As String = "smpti"
    Dim ArqLog As String = ""
    Dim From As String = "suporte@icraft.com.br"
    Dim Subject As String = "SYNCDOS - Sincronismo entre diretórios"
    Dim SemConfirm As Boolean = False
    Dim Help As Boolean = False
    Dim Result As String = ""

    Dim Replica As Icraft.IcftBase.DirReplica

    Sub Main()

        Dim Args As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs()
        Dim Tab As String = ""
        Dim Diretorio As String = ""
        Dim IncluiSubDir As Boolean = False


        For z As Integer = 0 To Args.Count - 1 Step 2
            Dim Coma As String = SemAspas(Args(z)).ToLower

            Select Case Coma
                Case "-recurso"
                    Recurso = SemAspas(Args(z + 1))
                Case "-email"
                    EmailPara = SemAspas(Args(z + 1))
                Case "-smtp"
                    Servidor = SemAspas(Args(z + 1))
                Case "-dir"
                    Diretorio = SemAspas(Args(z + 1))
                    IncluiSubDir = True
                Case "-dirsemsub"
                    Diretorio = SemAspas(Args(z + 1))
                    IncluiSubDir = False
                Case "-repl"
                    Tab &= IIf(Tab <> "", ";", "") & Diretorio & ";" & SemAspas(Args(z + 1)) & ";" & IIf(IncluiSubDir, "True", "False")
                Case "-arqlog"
                    ArqLog = SemAspas(Args(z + 1))
                Case "-from"
                    From = SemAspas(Args(z + 1))
                Case "-subject"
                    Subject = SemAspas(Args(z + 1))
                Case "-semconfirm"
                    SemConfirm = True
                    z -= 1
                Case "-help"
                    Help = True
            End Select
        Next

        If Help Or Args.Count = 0 Then
            Dim Msg As String = vbCrLf & vbCrLf & "SyncDOS - V04.00 - Programa de sincronização de arquivos e diretórios" & vbCrLf & vbCrLf
            Msg &= "     Modo de usar (exemplo)......................................." & vbCrLf
            Msg &= "     mostrar help:               -help" & vbCrLf
            Msg &= "     enviar email no final:      -email ""email@icraft.com.br""" & vbCrLf
            Msg &= "     utilizar smtp:              -smtp ""smtpi.icraft.com.br""" & vbCrLf
            Msg &= "     diretório incluindo subs:   -dir ""c:\origem""" & vbCrLf
            Msg &= "     diretório sem sub-dir:      -dirsemsub ""c:\origemsemsub""" & vbCrLf
            Msg &= "     diretório réplica:          -repl ""c:\destino""" & vbCrLf
            Msg &= "     gravar log em (arquivo):    -log ""c:\sync.log""" & vbCrLf
            Msg &= "     subject da mensagem:        -subject ""Sync Componentes""" & vbCrLf
            Msg &= "     from da mensagem:           -from ""'Suporte' [web@icraft.com.br]""" & vbCrLf
            Msg &= "     sem confirmar:              -semconfirm" & vbCrLf
            Msg &= "     recurso controlado:         -recurso ""bkp 10.0.0.70""" & vbCrLf
            Msg &= vbCrLf
            System.Console.WriteLine(Msg)
        End If

        If Tab = "" Then
            System.Console.WriteLine("[ERRO] Necessario definir lista de diretórios de origem e destino.")
            End
        End If

        If Not SemConfirm Then
            If MsgBox("Certeza de alterar, podendo até excluir, arquivos e diretórios REPLICADOS???", MsgBoxStyle.Critical + MsgBoxStyle.OkCancel + MsgBoxStyle.DefaultButton2) = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        End If

        Dim MomentoIni As Date = Now

        Dim StrCols As String = ""
        Dim Opc() As String = Split(Tab, ";")
        For z As Integer = 0 To Opc.Count - 1 Step 3
            Dim Origem As String = Opc(z)
            Dim Repl As String = Opc(z + 1)
            Dim IncluiSub As Boolean = Opc(z + 2)

            If Origem <> "" And Repl <> "" Then
                Replica = New Icraft.IcftBase.DirReplica(Origem, Repl, IncluiSub)
                AddHandler Replica.NotificaStatus, AddressOf Notifica
                Replica.LogDetalhado = True
                Replica.Executa()
                Result &= IIf(Result <> "", vbCrLf, "") & Replica.Log.ToString
            End If
        Next

        Dim MomentoFim As Date = Now
        Result &= "Recurso: " & Subject & vbCrLf & "Início:  " & Format(MomentoIni, "yyyy-MM-dd HH:mm:ss") & vbCrLf & "Término: " & Format(MomentoFim, "yyyy-MM-dd HH:mm:ss") & vbCrLf & "Duração: " & ExibeSegs(DateDiff(DateInterval.Second, MomentoIni, MomentoFim), ExibeSegsOpc.hh_mm_ss) & vbCrLf & vbCrLf & Result

        Dim STREMAIL As String = Trim(Icraft.IcftBase.EmailStr(EmailPara))
        If STREMAIL <> "" Then
            EnviaEmail(NZV(From, "<suporte@icraft.com.br>"), STREMAIL, "Sync - Log - " & Environ("COMPUTERNAME") & " - " & Environ("USERNAME") & IIf(Subject <> "", " - " & Subject, ""), "<div style='font-family:arial;font-size:8pt'>Resultado da sincronização:<ul><li>" & System.Text.RegularExpressions.Regex.Replace(Result.Trim(vbCrLf).Replace(vbCrLf, "</li><li>"), "(?is)(\[erro\]|\[falha\])", "<span style='background-color:yellow'>$1</span>") & "</li></ul></div>", , Servidor)
        End If

        If ArqLog <> "" Then
            If System.IO.File.Exists(ArqLog) Then
                Kill(ArqLog)
            End If
            Icraft.IcftBase.GravaLog(ArqLog, Result)
        End If

        If Not SemConfirm Then
            MsgBox(Replica.Status)
        End If
    End Sub

    Sub Notifica(Optional ByVal Texto As String = "")
        System.Console.WriteLine("> " & Format(Now, "yyyy-MM-dd HH:mm") & " - " & IIf(Texto <> "", "", Replica.Status))
    End Sub

End Module
