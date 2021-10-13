Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Net.Mail
Imports System.Drawing
Imports System.IO

Module Module1
    Dim Recurso As String = ""
    Dim Arquivos As String = ""
    Dim EmailPara As String = ""
    Dim Servidor As String = "smpti"
    Dim ArqLog As String = ""
    Dim From As String = "suporte@icraft.com.br"
    Dim Subject As String = "SYNCDOS - Sincronismo entre diretórios"
    Dim SemConfirm As Boolean = False
    Dim Help As Boolean = False
    Dim Result As String = ""

    Dim Replica As DirReplica

    Sub Main()
        Try
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
                    Case "-arquivos"
                        Arquivos = SemAspas(Args(z + 1))
                    Case "-help"
                        Help = True
                End Select
            Next

            If Help Or Args.Count = 0 Then
                Dim Msg As String = vbCrLf & vbCrLf & "SyncDOS - Programa de sincronização de arquivos e diretórios - " & VersaoApl() & vbCrLf & vbCrLf
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
                Msg &= "     arquivos:                   -arquivos ""\.aspx$|\.vb$"" (regex)"
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
                    Replica = New DirReplica(Origem, Repl, IncluiSub)
                    AddHandler Replica.NotificaStatus, AddressOf Notifica
                    Replica.LogDetalhado = True
                    Replica.Executa()
                    Result &= IIf(Result <> "", vbCrLf, "") & Replica.Log.ToString
                End If
            Next

            Dim MomentoFim As Date = Now
            Result &= "Recurso: " & Subject & vbCrLf & "Início:  " & Format(MomentoIni, "yyyy-MM-dd HH:mm:ss") & vbCrLf & "Término: " & Format(MomentoFim, "yyyy-MM-dd HH:mm:ss") & vbCrLf & "Duração: " & ExibeSegs(DateDiff(DateInterval.Second, MomentoIni, MomentoFim), ExibeSegsOpc.hh_mm_ss) & vbCrLf & vbCrLf & Result

            If EmailPara <> "" Then
                Dim STREMAIL As String = Trim(EmailStr(EmailPara))
                EnviaEmail(NZV(From, "<suporte@icraft.com.br>"), STREMAIL, "Sync - Log - " & Environ("COMPUTERNAME") & " - " & Environ("USERNAME") & IIf(Subject <> "", " - " & Subject, ""), "<div style='font-family:arial;font-size:8pt'>Resultado da sincronização:<ul><li>" & System.Text.RegularExpressions.Regex.Replace(Result.Trim(vbCrLf).Replace(vbCrLf, "</li><li>"), "(?is)(\[erro\]|\[falha\])", "<span style='background-color:yellow'>$1</span>") & "</li></ul></div>", , Servidor)
            End If

            If ArqLog <> "" Then
                If System.IO.File.Exists(ArqLog) Then
                    Kill(ArqLog)
                End If
                GravaLog(ArqLog, Result)
            End If

            If Not SemConfirm Then
                MsgBox(Replica.Status)
            End If
        Catch EX As Exception
            System.Console.WriteLine("Erro: " & EX.Message)
        End Try
    End Sub

    Public Class DirReplica

        ''' <summary>
        ''' Cria instrução para replicação de diretórios.
        ''' </summary>
        ''' <param name="DirOrigem">Diretório de origem que permanecerá inalterado.</param>
        ''' <param name="DirDestino">Diretório de destino, que será alterado.</param>
        ''' <param name="IncluirSubDir">Incluir sub-diretórios?</param>
        ''' <param name="ApagarQuandoEncontrar">Apagar quais arquivos quando encontrar (thumbmail por exemplo)?</param>
        ''' <remarks></remarks>
        Sub New(ByVal DirOrigem As String, ByVal DirDestino As String, ByVal IncluirSubDir As Boolean, Optional ByVal ApagarQuandoEncontrar As String = "")
            Me.DirOrigem = DirOrigem
            Me.DirDestino = DirDestino
            Me.IncluiSub = IncluirSubDir
            Me._ApagarQuandoEncontrar = ApagarQuandoEncontrar
        End Sub
        Private _ApagarQuandoEncontrar As String
        Private _dirorigem As String

        ''' <summary>
        ''' Diretório de origem.
        ''' </summary>
        ''' <value>Diretório de origem.</value>
        ''' <returns>Diretório de origem.</returns>
        ''' <remarks></remarks>
        Public Property DirOrigem() As String
            Get
                Return _dirorigem
            End Get
            Set(ByVal value As String)
                _dirorigem = value
            End Set
        End Property

        Private _dirdestino As String

        ''' <summary>
        ''' Diretório de destino.
        ''' </summary>
        ''' <value>Diretório de destino.</value>
        ''' <returns>Diretório de destino.</returns>
        ''' <remarks></remarks>
        Public Property DirDestino() As String
            Get
                Return _dirdestino
            End Get
            Set(ByVal value As String)
                _dirdestino = value
            End Set
        End Property

        Private _incluisub As Boolean

        ''' <summary>
        ''' Inclui sub-diretório.
        ''' </summary>
        ''' <value>Inclui sub-diretório.</value>
        ''' <returns>Inclui sub-diretório.</returns>
        ''' <remarks></remarks>
        Public Property IncluiSub() As Boolean
            Get
                Return _incluisub
            End Get
            Set(ByVal value As Boolean)
                _incluisub = value
            End Set
        End Property

        Private _qtdarqs As Integer

        ''' <summary>
        ''' Quantidade de arquivos.
        ''' </summary>
        ''' <value>Quantidade de arquivos.</value>
        ''' <returns>Quantidade de arquivos.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property QtdArqs() As Integer
            Get
                Return _qtdarqs
            End Get
        End Property

        ''' <summary>
        ''' Executa rotina de notificação de status externa (carregada por delegate).
        ''' </summary>
        ''' <remarks></remarks>
        Public Event NotificaStatus()

        Private _status As String

        ''' <summary>
        ''' Status do sincronismo.
        ''' </summary>
        ''' <value>Status do sincronismo.</value>
        ''' <returns>Status do sincronismo.</returns>
        ''' <remarks></remarks>
        Public Property Status() As String
            Get
                Return _status
            End Get
            Set(ByVal value As String)
                _status = value
            End Set
        End Property

        Private _inicio As Date = Nothing

        ''' <summary>
        ''' Início da replicação.
        ''' </summary>
        ''' <value>Início da replicação.</value>
        ''' <returns>Início da replicação.</returns>
        ''' <remarks></remarks>
        Public Property Inicio() As Date
            Get
                Return _inicio
            End Get
            Set(ByVal value As Date)
                _inicio = value
            End Set
        End Property

        Private _termino As Date = Nothing

        ''' <summary>
        ''' Término da replicação.
        ''' </summary>
        ''' <value>Término da replicação.</value>
        ''' <returns>Término da replicação.</returns>
        ''' <remarks></remarks>
        Public Property Termino() As Date
            Get
                Return _termino
            End Get
            Set(ByVal value As Date)
                _termino = value
            End Set
        End Property

        Private _logdetalhado As Boolean = False

        ''' <summary>
        ''' Registro de log detalhado.
        ''' </summary>
        ''' <value>Registro de log detalhado.</value>
        ''' <returns>Registro de log detalhado.</returns>
        ''' <remarks></remarks>
        Public Property LogDetalhado() As Boolean
            Get
                Return _logdetalhado
            End Get
            Set(ByVal value As Boolean)
                _logdetalhado = value
            End Set
        End Property

        Private _log As New StringBuilder

        ''' <summary>
        ''' Registro de log.
        ''' </summary>
        ''' <value>Registro de log.</value>
        ''' <returns>Registro de log.</returns>
        ''' <remarks></remarks>
        Public Property Log() As StringBuilder
            Get
                Return _log
            End Get
            Set(ByVal value As StringBuilder)
                _log = value
            End Set
        End Property

        ''' <summary>
        ''' Execução da replicação de arquivo.
        ''' </summary>
        ''' <param name="Arquivo">Arquivo a ser replicado.</param>
        ''' <param name="DirOrigem">Diretório de origem.</param>
        ''' <param name="DirDestino">Diretório de destino.</param>
        ''' <param name="ArquivoDest">Arquivo de destino.</param>
        ''' <remarks></remarks>
        Private Sub Trata(ByVal Arquivo As String, ByVal DirOrigem As String, ByVal DirDestino As String, Optional ByVal ArquivoDest As String = "")
            Try
                If ArquivoDest = "" Then
                    ArquivoDest = Arquivo
                End If

                Dim ArqO As String = FileExpr(DirOrigem, Arquivo)
                Dim ArqD As String = FileExpr(DirDestino, ArquivoDest)


                Dim Tratou As Boolean = False
                If _ListaApagar.Length > 0 Then
                    For Each Item As String In _ListaApagar
                        Try
                            If Item <> "" AndAlso Arquivo Like Item Then
                                Try
                                    If System.IO.File.Exists(ArqO) Then
                                        System.IO.File.SetAttributes(ArqO, FileAttributes.Archive)
                                        System.IO.File.Delete(ArqO)
                                        RegLog("Eliminado arquivo " & ArqO & " por corresponder à máscara '" & Item & "'")
                                    End If
                                Catch ex As Exception
                                    RegLog("[FALHA] " & ex.Message & " ao tentar excluir arquivo " & ArqO)
                                End Try
                                Try
                                    If System.IO.File.Exists(ArqD) Then
                                        System.IO.File.SetAttributes(ArqD, FileAttributes.Archive)
                                        System.IO.File.Delete(ArqD)
                                        RegLog("Eliminado arquivo " & ArqD & " por corresponder à máscara '" & Item & "'")
                                    End If
                                Catch ex As Exception
                                    RegLog("[FALHA] " & ex.Message & " ao tentar excluir arquivo " & ArqD)
                                End Try
                                Tratou = True
                                Exit For
                            End If
                        Catch ex As Exception
                            RegLog("[FALHA] " & ex.Message & " ao tentar validar expressão " & Item & " para arquivo " & Arquivo)
                        End Try
                    Next
                End If

                If Not Tratou Then
                    Dim AtribOrigem As New System.IO.FileInfo(ArqO)
                    Dim AtribDestino As New System.IO.FileInfo(ArqD)
                    If AtribOrigem.Exists Then
                        Dim TempoDif As Boolean = False
                        Try
                            TempoDif = AtribOrigem.LastWriteTime <> AtribDestino.LastWriteTime
                        Catch ex As Exception
                            RegLog("[FALHA] " & ex.Message & " ao tentar obter lastwritetime de " & ArqD)
                        End Try

                        If (Not AtribDestino.Exists) OrElse TempoDif OrElse (AtribOrigem.Length <> AtribDestino.Length) Then
                            Try
                                FileCopy(ArqO, ArqD)
                            Catch EX As System.IO.DirectoryNotFoundException
                                Try
                                    CriaDir(DirDestino)
                                    FileCopy(ArqO, ArqD)
                                Catch Ex2 As Exception
                                    RegLog("[FALHA] " & Ex2.Message & " ao tentar copiar " & ArqO & " para " & ArqD)
                                End Try
                            Catch ex As Exception
                                RegLog("[FALHA] " & ex.Message & " ao tentar copiar " & ArqO & " para " & ArqD)
                            End Try
                        End If
                    End If
                End If
            Catch ex As Exception
                RegLog("[FALHA] " & ex.Message & " ao tentar copiar " & Arquivo & " para " & ArquivoDest)
            End Try
        End Sub

        ''' <summary>
        ''' Registro de log de sincronização.
        ''' </summary>
        ''' <param name="Texto">Mensagem que será registrada.</param>
        ''' <remarks></remarks>
        Private Sub RegLog(ByVal Texto As String)
            Log.AppendLine(Format(Now, "ddd HH:mm:ss") & " - " & Texto)
        End Sub

        ''' <summary>
        ''' Apaga arquivo.
        ''' </summary>
        ''' <param name="Arquivo">Arquivo que será apagado.</param>
        ''' <param name="Diretorio">Diretório onde se encontra este arquivo.</param>
        ''' <remarks></remarks>
        Private Sub Apaga(ByVal Arquivo As String, ByVal Diretorio As String)
            Try
                Dim Arq As String = FileExpr(Diretorio, Arquivo)
                Kill(Arq)
                If LogDetalhado Then
                    RegLog("Apagou " & Arq)
                End If
            Catch Ex As Exception
                RegLog("[FALHA] " & Ex.Message & " ao tentar excluir " & Arquivo & " do diretório " & Diretorio)
            End Try
        End Sub

        ''' <summary>
        ''' Criação de diretório.
        ''' </summary>
        ''' <param name="Diretorio">Diretório que será criado.</param>
        ''' <remarks></remarks>
        Private Sub CriaDir(ByVal Diretorio As String)
            Try
                MkDir(Diretorio)
                If LogDetalhado Then
                    RegLog("Criou " & Diretorio)
                End If
            Catch ex As Exception
                RegLog("[FALHA] " & ex.Message & " ao tentar criar diretório " & Diretorio)
            End Try
        End Sub

        Dim UltNotif As String = ""

        ''' <summary>
        ''' Registro de notificação de status.
        ''' </summary>
        ''' <param name="Texto">Texto a notificar.</param>
        ''' <param name="Forcar">Grava mensagem e manda para função de notificação agora.</param>
        ''' <remarks></remarks>
        Private Sub Notifica(ByVal Texto As String, Optional ByVal Forcar As Boolean = False)
            Dim Notif As String = Format(Now, "ss")
            If UltNotif <> Notif Or Forcar Then
                Status = Texto
                UltNotif = Notif
                RaiseEvent NotificaStatus()
            End If
        End Sub

        ''' <summary>
        ''' Execução do processo de replicação.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Executa()
            Try
                Inicio = Now
                Dim Ocorr As String = "Início de replicação de " & DirOrigem & " para " & DirDestino
                Notifica(Ocorr, True)
                RegLog(Ocorr)

                Executa(DirOrigem, DirDestino, Arquivos)
                Termino = Now

                Ocorr = "Término (" & QtdArqs & Pl(QtdArqs, " arquivo") & " | duração: " & ExibeSegs(DateDiff(DateInterval.Second, Inicio, Termino), ExibeSegsOpc.xh_ymin_zseg) & ")"
                Notifica(Ocorr, True)
                RegLog(Ocorr)
            Catch ex As Exception
                RegLog("[FALHA] " & ex.Message & " ao tentar executar sincronização entre " & DirOrigem & " e " & DirDestino)
            End Try
        End Sub

        Private DirBloqueado() As String = {"$RECYCLE.BIN", "System Volume Information"}

        ''' <summary>
        ''' Verifica se caminho está bloqueado.
        ''' </summary>
        ''' <param name="Caminho">Caminho.</param>
        ''' <returns>True caso esteja na lista de caminhos bloqueados ou false caso contrário.</returns>
        ''' <remarks></remarks>
        Private Function Bloqueado(ByVal Caminho As String) As Boolean
            Try
                Dim Disco As String = System.IO.Path.GetPathRoot(Caminho)
                Caminho = Mid(Caminho, Len(Disco) + 1) & "\"
                For Each Item As String In DirBloqueado
                    If Caminho.StartsWith(Item & "\") Then
                        Return True
                    End If
                Next
            Catch ex As Exception
                RegLog("[FALHA] " & ex.Message & " ao verificar se caminho " & Caminho & " está na lista de itens bloqueados")
            End Try
            Return False
        End Function

        ''' <summary>
        ''' Garante gravação adequada do arquivo.
        ''' </summary>
        ''' <param name="Arquivo">Arquivo.</param>
        ''' <param name="DirOrigem">Diretório de origem.</param>
        ''' <param name="DirDestino">Diretório de destino.</param>
        ''' <param name="ArquivoDest">Arquivo de destino.</param>
        ''' <remarks></remarks>
        Private Sub Garante(ByVal Arquivo As String, ByVal DirOrigem As String, ByVal DirDestino As String, Optional ByVal ArquivoDest As String = "")
            Try
                If ArquivoDest = "" Then
                    ArquivoDest = Arquivo
                End If

                If Not System.IO.File.Exists(FileExpr(DirOrigem, Arquivo)) AndAlso System.IO.File.Exists(FileExpr(DirDestino, Arquivo)) Then
                    Apaga(ArquivoDest, DirDestino)
                End If
            Catch ex As Exception
                RegLog("[FALHA] " & ex.Message & " ao buscar garantias de igualdade entre origem " & DirOrigem & "..." & Arquivo & " e " & DirDestino & "..." & ArquivoDest)
            End Try
        End Sub

        Private _ListaApagar() As String = {}

        ''' <summary>
        ''' Executa replicação de diretório de origem para réplica.
        ''' </summary>
        ''' <param name="Origem">Diretório de origem.</param>
        ''' <param name="Destino">Diretório de réplica.</param>
        ''' <remarks></remarks>
        Private Sub Executa(ByVal Origem As String, ByVal Destino As String, Arquivos As String)
            Try
                _ListaApagar = Split(_ApagarQuandoEncontrar, vbCrLf)
                If _ListaApagar.Length = 1 AndAlso Trim(_ListaApagar(0)) = "" Then
                    _ListaApagar = New String() {}
                End If

                ' garante todos os arquivos da origem no destino
                If System.IO.Directory.Exists(Origem) Then


                    For Each Arq As String In System.IO.Directory.GetFiles(Origem)
                        If Arquivos <> "" AndAlso Not System.Text.RegularExpressions.Regex.IsMatch(Arq, Arquivos) Then Continue For
                        If Not Bloqueado(Arq) Then
                            Notifica(Origem)

                            Dim ArqA As String = System.IO.Path.GetFileName(Arq)
                            Trata(ArqA, Origem, Destino)
                            _qtdarqs += 1
                        End If
                    Next

                    ' garante que não tenha nenhum a mais
                    If System.IO.Directory.Exists(Destino) Then
                        For Each Arq As String In System.IO.Directory.GetFiles(Destino)
                            If Arquivos <> "" AndAlso Not System.Text.RegularExpressions.Regex.IsMatch(Arq, Arquivos) Then Continue For
                            If Not Bloqueado(Arq) Then
                                Notifica(Arq)

                                Dim ArqA As String = System.IO.Path.GetFileName(Arq)
                                Garante(ArqA, Origem, Destino)
                            End If
                        Next
                    Else
                        CriaDir(Destino)
                    End If

                    ' diretório que existem na origem
                    For Each Dir As String In System.IO.Directory.GetDirectories(Origem)
                        If Not Bloqueado(Dir) Then
                            Notifica(Dir)

                            Dim DirA As String = System.IO.Path.GetFileName(Dir)
                            If IncluiSub Then
                                Executa(Dir, FileExpr(Destino, DirA), Arquivos)
                            End If
                        End If
                    Next

                    ' diretórios existentes no destino sem origem
                    For Each Dir As String In System.IO.Directory.GetDirectories(Destino)
                        If Not Bloqueado(Dir) Then
                            Notifica(Dir)

                            Dim DirA As String = System.IO.Path.GetFileName(Dir)
                            If Not System.IO.Directory.Exists(FileExpr(Origem, DirA)) Then
                                ApagaDir(Dir)
                            End If
                        End If
                    Next
                Else
                    Dim OrigArq As String = System.IO.Path.GetFileName(Origem)
                    If OrigArq <> "" Then
                        Dim OrigSemArq As String = System.IO.Path.GetDirectoryName(Origem)
                        Dim DestArq As String = System.IO.Path.GetFileName(Destino)
                        Dim DestSemArq As String = System.IO.Path.GetDirectoryName(Destino)

                        Trata(OrigArq, OrigSemArq, DestSemArq, DestArq)
                        _qtdarqs += 1
                        Garante(OrigArq, OrigSemArq, DestSemArq, DestArq)

                    End If
                End If
            Catch ex As Exception
                RegLog("[FALHA] " & ex.Message & " ao executar sincronização entre origem " & Origem & " e destino " & Destino)
            End Try
        End Sub

        ''' <summary>
        ''' Apagar diretório.
        ''' </summary>
        ''' <param name="Diretorio">Diretório a ser eliminado.</param>
        ''' <remarks></remarks>
        Sub ApagaDir(ByVal Diretorio As String)
            Try
                System.IO.Directory.Delete(Diretorio, True)
                If LogDetalhado Then
                    RegLog("Apagou " & Diretorio)
                End If
            Catch EX As Exception
                RegLog("[FALHA] ao apagar diretório " & Diretorio & ": " & EX.Message)
            End Try
        End Sub

    End Class

    Function FileExpr(ByVal ParamArray Segmentos() As String) As String
        Dim Raiz As String = New System.Web.UI.Control().ResolveUrl("~/").Replace("/", "\")
        Dim Arq As String = ExprExpr("\", "/", "", Segmentos)
        If Arq.StartsWith(Raiz) Then
            Arq = "~\" & Mid(Arq, Len(Raiz) + 1)
        End If

        If Arq.StartsWith("~\") Then
            If Ambiente() = AmbienteTipo.WEB Then
                Arq = System.Web.HttpContext.Current.Server.MapPath(Arq)
            Else
                Dim DirExec As String = FileExpr(WebConf("dir_raiz_site"), "\")
                If DirExec = "" Or DirExec = "\" Then
                    DirExec = System.Reflection.Assembly.GetExecutingAssembly().Location
                End If
                Arq = Arq.Replace("~\", System.IO.Path.GetDirectoryName(DirExec) & "\")
            End If
        End If
        Return Arq
    End Function

    Function ExprExpr(ByVal Delim As String, ByVal DelimAlternativo As String, ByVal Inicial As Object, ByVal ParamArray Segmentos() As Object) As String
        Inicial = NZ(Inicial, "")
        Dim Lista As ArrayList = ParamArrayToArrayList(Segmentos)
        For Each item As Object In Lista
            If Not IsNothing(item) Then
                If Not IsNothing(DelimAlternativo) AndAlso DelimAlternativo <> "" Then
                    item = item.Replace(DelimAlternativo, Delim)
                End If
                item = NZ(item, "")
                If item <> "" Then
                    If Inicial <> "" Then
                        If Inicial.EndsWith(Delim) AndAlso item.StartsWith(Delim) Then
                            Inicial &= CType(item, String).Substring(Delim.Length)
                        ElseIf Inicial.EndsWith(Delim) OrElse item.StartsWith(Delim) Then
                            Inicial &= item
                        Else
                            Inicial &= Delim & item
                        End If
                    Else
                        Inicial &= item
                    End If
                End If
            End If
        Next
        Return Inicial
    End Function

    Function NZ(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
        Dim tipo As String

        If Not IsNothing(Def) Then
            tipo = Def.GetType.ToString
        ElseIf IsNothing(Valor) Then
            Return Nothing
        Else
            tipo = Valor.GetType.ToString.Trim
        End If

        If IsNothing(Valor) OrElse IsDBNull(Valor) OrElse ((tipo = "System.DateTime" Or Valor.GetType.ToString = "System.DateTime") AndAlso Valor = CDate(Nothing)) Then
            Valor = Def
        End If

        Select Case tipo
            Case "System.Decimal"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Decimal)
                End If
                Return CType(Valor, Decimal)
            Case "System.String"
                If Valor.GetType.ToString = "System.Byte[]" Then
                    Return CType(ByteArrayToObject(Valor), String)
                End If
                If Valor.GetType.ToString = "Icraft.IcftBase+LogonSession" Then
                    Return CType(Valor, LogonSession).ToString
                ElseIf Valor.GetType.IsEnum Then
                    Return Valor.ToString
                End If
                Return CType(Valor, String)
            Case "System.Double"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Double)
                End If
                Return CType(Valor, Double)
            Case "System.Boolean"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return False
                End If
                Return CType(Valor, Boolean)
            Case "System.DateTime"
                Return CType(Valor, System.DateTime)
            Case "System.Single"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Single)
                End If
                Return CType(Valor, System.Single)
            Case "System.Byte"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Byte)
                End If
                Return CType(Valor, System.Byte)
            Case "System.Char"
                Return CType(Valor, System.Char)
            Case "System.SByte"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, SByte)
                End If
                Return CType(Valor, System.SByte)
            Case "System.Int32"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Int32)
                End If
                Return CType(Valor, Int32)
            Case "System.DBNull"
                Return Valor
            Case "System.Collections.ArrayList"
                Return ParamArrayToArrayList(Valor)
            Case "System.Data.DataSet"
                If IsNothing(Valor) Then
                    Return Def
                End If
                Return Valor
        End Select

        Return CType(Valor, String)
    End Function


    Function ParamArrayToArrayList(ByVal ParamArray Params() As Object) As Object

        ' caso não existam parâmetros
        If IsNothing(Params) OrElse Params.Length = 0 Then
            Return New ArrayList
        End If

        ' caso já seja um arraylist
        If Params.Length = 1 And TypeOf (Params(0)) Is ArrayList Then
            Return Params(0)
        End If

        ' caso tenha que juntar
        Dim ListaParametros As ArrayList = New ArrayList
        For Each Item As Object In Params
            If Not IsNothing(Item) Then

                ' >> TIPOS PREVISTOS EM ARRAYLIST...
                ' array
                ' arraylist
                ' string
                ' dataset
                ' datarowcollection

                If TypeOf Item Is Array Then
                    For Each SubItem As Object In Item
                        ListaParametros.AddRange(ParamArrayToArrayList(SubItem))
                    Next
                ElseIf TypeOf Item Is ArrayList OrElse Item.GetType.ToString.StartsWith("System.Collections.Generic.List") Then
                    ListaParametros.AddRange(Item)
                ElseIf TypeOf Item Is String Then
                    ListaParametros.Add(Item)
                ElseIf TypeOf Item Is DataSet Then
                    For Each Row As DataRow In Item.Tables(0).rows
                        For Each Campo As Object In Row.ItemArray
                            ListaParametros.Add(Campo)
                        Next
                    Next
                ElseIf TypeOf Item Is DataRow Then
                    For Each Campo As Object In CType(Item, DataRow).ItemArray
                        ListaParametros.Add(Campo)
                    Next
                ElseIf TypeOf Item Is System.IO.FileInfo Then
                    ListaParametros.Add(Item.name)
                Else
                    ListaParametros.Add(Item)
                End If
            End If
        Next
        Return ListaParametros
    End Function

    Public Function Pl(ByVal Numero As Object, ByVal Singular As String, Optional ByVal Plural As String = "") As String
        Return IIf(Numero = 1, Singular, NZV(Plural, Singular & IIf(Char.IsLower(Microsoft.VisualBasic.Right(Singular, 1)), "s", "S")))
    End Function

    Public Function ExibeSegs(ByVal QtdSegundos As Integer, ByVal Opc As ExibeSegsOpc) As String
        Dim Segs As Integer = QtdSegundos
        Dim Horas As Integer = Int(Segs / 3600)
        Segs -= Horas * 3600
        Dim Mins As Integer = Int(Segs / 60)
        Segs -= Mins * 60

        Select Case Opc
            Case ExibeSegsOpc.hh_mm_ss
                Return Format(Horas, "00") & ":" & Format(Mins, "00") & ":" & Format(Segs, "00")
            Case ExibeSegsOpc.x_horas_y_minutos_e_z_segundos
                Return Horas & Pl(Horas, " Hora") & ", " & Mins & Pl(Mins, " Minuto") & " e " & Segs & Pl(Segs, " Segundo")
            Case ExibeSegsOpc.xh_ymin_zseg
                Return Horas & "h " & Mins & "min " & Segs & "seg"
            Case ExibeSegsOpc.hh_mm
                Return Format(Horas, "00") & ":" & Format(Mins, "00")
            Case ExibeSegsOpc.d_dias_x_horas_y_minutos_e_z_segundos

                Dim dia As Integer = Int(Horas / 24)
                Dim hora As Integer = (Horas Mod 24)

                Return dia & Pl(dia, " Dia") & ", " & hora & Pl(hora, " Hora") & ", " & Mins & Pl(Mins, " Minuto") & " e " & Segs & Pl(Segs, " Segundo")

            Case ExibeSegsOpc.mm_ss
                Dim seg As Integer = QtdSegundos
                Dim Min As Integer = (seg / 60)
                Return Min & " min" & ", " & Segs & " seg"
        End Select
        Return QtdSegundos & Pl(QtdSegundos, " Segundo")

    End Function

    Public Enum ExibeSegsOpc
        xh_ymin_zseg
        hh_mm_ss
        x_horas_y_minutos_e_z_segundos
        z_segundos
        hh_mm
        d_dias_x_horas_y_minutos_e_z_segundos
        mm_ss
    End Enum

    Public Function Ambiente() As AmbienteTipo
        Try
            If Not IsNothing(System.Web.HttpContext.Current) Then
                Return AmbienteTipo.WEB
            End If
        Catch
        End Try
        Return AmbienteTipo.Windowsforms
    End Function

    Public Enum AmbienteTipo
        Windowsforms
        WEB
    End Enum

    Function WebConf(ByVal param As String) As String
        If Compare(param, "SITE_DIR") Then
            Return FileExpr("~/")
        ElseIf Compare(param, "SITE_URL") Then
            Return URLExpr("~/")
        End If
        Return System.Configuration.ConfigurationManager.AppSettings(param)
    End Function

    Function Compare(ByVal Param1 As Object, ByVal Param2 As Object, Optional ByVal IgnoreCase As Boolean = True) As Boolean
        If IsNothing(Param1) And IsNothing(Param2) Then
            Return True
        ElseIf IsNothing(Param1) Or IsNothing(Param2) Then
            Return False
        Else
            If Param1.GetType.ToString = Param2.GetType.ToString Then
                If Param1.GetType.ToString = "System.String" Then
                    Return String.Compare(Param1, Param2, IgnoreCase) = 0
                Else
                    Err.Raise(20000, "IcraftBase", "Compare com tipo não previsto " & Param1.GetType.ToString & ".")
                End If
            End If
        End If
        Return False
    End Function

    Function SqlExpr(ByVal Conteudo As Object, Optional ByVal CaracAbreFechaString As String = "'") As String
        If TypeOf (Conteudo) Is String Then
            Return CaracAbreFechaString & Replace(Conteudo, CaracAbreFechaString, CaracAbreFechaString & CaracAbreFechaString) & CaracAbreFechaString
        ElseIf TypeOf (Conteudo) Is DBNull Then
            Return "NULL"
        ElseIf TypeOf Conteudo Is Decimal OrElse TypeOf Conteudo Is Double OrElse TypeOf Conteudo Is Single OrElse TypeOf Conteudo Is Int32 OrElse TypeOf Conteudo Is Byte Then
            Return Str(Conteudo)
        ElseIf TypeOf (Conteudo) Is Boolean Then
            Return IIf(Conteudo, Boolean.TrueString, Boolean.FalseString)
        ElseIf TypeOf (Conteudo) Is Date Then
            Return "#" & Format(Conteudo, "yyyy-MM-dd HH:mm:ss") & "#"
        Else
            Throw New Exception("Tipo desconhecido " & Conteudo.GetType.ToString & " para obtenção de expressão para sql.")
        End If
    End Function

    Function NZV(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
        Dim Result As Object = NZ(Valor, Def)
        If TypeOf Result Is String AndAlso Result = "" Then
            Return Def
        ElseIf TypeOf Result Is Decimal AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Double AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Single AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Int32 AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Byte AndAlso Result = 0 Then
            Return Def
        End If
        Return Result
    End Function

    Function ByteArrayToObject(ByVal Bytes() As Byte) As Object
        Dim Obj As Object = Nothing
        Try
            Dim fs As System.IO.MemoryStream = New System.IO.MemoryStream
            Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
            fs.Write(Bytes, 0, Bytes.Length)
            fs.Seek(0, IO.SeekOrigin.Begin)

            Obj = formatter.Deserialize(fs)
        Catch
        End Try
        Return Obj
    End Function

    Function URLExpr(ByVal ParamArray Segmentos() As Object) As String
        Dim URL As String = ExprExpr("/", "\", "", Segmentos)
        If Regex.Match(URL, "(?is)^[a-z0-9]:/").Success Then
            If Ambiente() = AmbienteTipo.WEB Then
                URL = URL.ToLower.Replace(System.Web.HttpContext.Current.Server.MapPath("~/").Replace("\", "/").ToLower, "~/")
            Else
                URL = URL.Replace("\", "/").ToLower
                URL = URL.Replace(FileExpr("~/").Replace("\", "/").ToLower, "~/")
            End If
        End If
        Return URL
    End Function

    Public Class LogonSession
        Private _id As String = Nothing
        Private _usuario As String = Nothing
        Private _momento As Date = Nothing
        Private _site As String = Nothing
        Private _senha As String = Nothing
        Private _grupo As String = Nothing
        Private _outros As New ArrayList

        ''' <summary>
        ''' Converte as informações de login da seção para uma string.
        ''' </summary>
        ''' <returns>Retorna a string contendo as informações.</returns>
        ''' <remarks></remarks>
        Public Shadows Function ToString() As String
            Dim txt As New StringBuilder
            txt.Append("LogonSession(")
            txt.Append("id=" & NZ(_id, "") & ";")
            txt.Append("_usuario=" & NZ(_usuario, "") & ";")
            txt.Append("_momento=" & Format(NZV(_momento, Nothing), "dd/MM/yyyy HH:mm:ss") & ";")
            For z As Integer = 0 To _outros.Count - 1 Step 2
                txt.Append(_outros(z) & "=")
                txt.Append(NZ(_outros(z + 1), ""))
                txt.Append(";")
            Next
            txt.Append("_site=" & NZ(_site, ""))
            txt.Append("_grupo=" & NZ(_grupo, ""))
            txt.Append(")")
            Return txt.ToString
        End Function


        ''' <summary>
        ''' Identificação para armazenamento de logon do tipo 'GERAL' ou algum específico para múltiplos logons.
        ''' </summary>
        ''' <value>Especificação do tipo de logon.</value>
        ''' <returns>Especificação do tipo de logon.</returns>
        ''' <remarks></remarks>
        Public Property Id() As String
            Get
                Return _id
            End Get
            Set(ByVal value As String)
                _id = value
            End Set
        End Property

        ''' <summary>
        ''' Usuário que efetuou logon.
        ''' </summary>
        ''' <value>Login do usuário que efetuou logon.</value>
        ''' <returns>Login do usuário que efetuou logon.</returns>
        ''' <remarks></remarks>
        Public Property Usuario() As String
            Get
                Return _usuario
            End Get
            Set(ByVal value As String)
                _usuario = value
            End Set
        End Property

        ''' <summary>
        ''' Momento de logon.
        ''' </summary>
        ''' <value>Momento (data e hora) de logon.</value>
        ''' <returns>Momento (data e hora) de logon.</returns>
        ''' <remarks></remarks>
        Public Property Momento() As Date
            Get
                Return _momento
            End Get
            Set(ByVal value As Date)
                _momento = value
            End Set
        End Property

        ''' <summary>
        ''' Grupo do usuário que efetuou logon.
        ''' </summary>
        ''' <value>Nome do grupo do usuário que efetuou logon.</value>
        ''' <returns>Nome do grupo do usuário que efetuou logon.</returns>
        ''' <remarks></remarks>
        Public Property Grupo() As String
            Get
                Return _grupo
            End Get
            Set(ByVal value As String)
                _grupo = value
            End Set
        End Property

        ''' <summary>
        ''' Nome do site.
        ''' </summary>
        ''' <value></value>
        ''' <returns>Nome do site.</returns>
        ''' <remarks>Nome do site.</remarks>
        Public Property Site() As String
            Get
                Return _site
            End Get
            Set(ByVal value As String)
                _site = value
            End Set
        End Property

        ''' <summary>
        ''' Senha de acesso.
        ''' </summary>
        ''' <value>Senha de acesso.</value>
        ''' <returns>Senha de acesso.</returns>
        ''' <remarks></remarks>
        Public Property Senha() As String
            Get
                Return _senha
            End Get
            Set(ByVal value As String)
                _senha = value
            End Set
        End Property

        ''' <summary>
        ''' Outras propriedades a serem armazenadas pelo Logon.
        ''' </summary>
        ''' <param name="Propriedade">Nome da propriedade.</param>
        ''' <value>Valor da propriedade.</value>
        ''' <returns>Valor da propriedade armazenada.</returns>
        ''' <remarks></remarks>
        Public Property ExtendedProps(ByVal Propriedade As String) As Object
            Get
                Dim Pos As Integer = _outros.IndexOf(":" & Propriedade)
                If Pos >= 0 Then
                    Return _outros(Pos + 1)
                End If
                Return Nothing
            End Get
            Set(ByVal value As Object)
                Dim Pos As Integer = _outros.IndexOf(":" & Propriedade)
                If Pos >= 0 Then
                    _outros(Pos + 1) = value
                    Exit Property
                End If
                _outros.Add(":" & Propriedade)
                _outros.Add(value)
            End Set
        End Property

        ''' <summary>
        ''' Acesso aos atributos e propriedades expandidas.
        ''' </summary>
        ''' <param name="Nome">Nome da propriedade tratada.</param>
        ''' <value>Valor da propriedade tratada.</value>
        ''' <returns>Valor da propriedade solicitada.</returns>
        ''' <remarks></remarks>
        Default Property Attributes(ByVal Nome As String) As String
            Get
                If Compare(Nome, "Id") Then
                    Return _id
                ElseIf Compare(Nome, "Usuario") Then
                    Return _usuario
                ElseIf Compare(Nome, "Momento") Then
                    Return _momento
                ElseIf Compare(Nome, "Site") Then
                    Return _site
                ElseIf Compare(Nome, "Senha") Then
                    Return _senha
                ElseIf Compare(Nome, "Grupo") Then
                    Return _grupo
                Else
                    Dim Prop As Object = ExtendedProps(Nome)
                    If IsNothing(Prop) Then
                        Throw New Exception("Em Attributes de Logon, atributo '" & Nome & "' inválido para objeto " & Me.GetType.ToString & ".")
                    Else
                        Return Prop
                    End If
                End If
                Return Nothing
            End Get

            Set(ByVal value As String)
                If Compare(Nome, "Id") Then
                    _id = value
                ElseIf Compare(Nome, "Usuario") Then
                    _usuario = value
                ElseIf Compare(Nome, "Momento") Then
                    _momento = value
                ElseIf Compare(Nome, "Site") Then
                    _site = value
                ElseIf Compare(Nome, "Senha") Then
                    _senha = value
                ElseIf Compare(Nome, "Grupo") Then
                    _grupo = value
                Else
                    Throw New Exception("Em Attributes de Logon, atributo " & value & " inválido para objeto " & Me.GetType.ToString & ".")
                End If
            End Set
        End Property

        ''' <summary>
        ''' Criação de login para registro de acesso de usuário.
        ''' </summary>
        ''' <param name="Pagina">Página na qual é efetuado o login.</param>
        ''' <param name="Usuario">Usuário que efetua acesso.</param>
        ''' <param name="Senha">Senha do usuário.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Pagina As System.Web.UI.Page, ByVal Usuario As String, ByVal Senha As String)
            ' cria chave com area e usuario
            Try
                _id = Pagina.Session.SessionID
                _usuario = Usuario
                _momento = Now
                _site = WebConf("site_nome")
                _senha = Senha
            Catch
                _id = Nothing
                _usuario = Nothing
                _momento = Nothing
                _site = Nothing
                _senha = Nothing
            End Try
        End Sub

    End Class

    Function SemAspas(ByVal Texto As String) As String
        Return Texto.Trim("""", Chr(147), Chr(148))
    End Function

    Public Function EmailStr(ByVal Email As String) As String
        Email = Trim(Email)
        Email = Email.Replace("[", "<").Replace("]", ">").Replace(Chr(160), " ")
        If Email.StartsWith("'") Then
            Email = Regex.Replace(Email, "'(.*)'", """$1""")
        End If

        Email = Email.Replace("'", "`")

        Dim SoEmail As String = ""
        If Email.IndexOf("<") = -1 Then
            SoEmail = SoEmailStr(Email)
            If SoEmail <> "" Then
                Email = ReplRepl(Email, SoEmail, "")
            End If
        Else
            SoEmail = Regex.Match(Email, "<(.*?)>").Groups(1).Value
            Email = ReplRepl(Email, "<" & SoEmail & ">", "")
        End If

        Email = ReplRepl(Email, Chr(9), "")
        Email = TrimCarac(Trim(ReplRepl(Email, "  ", " ")), New String() {Chr(34), "'"})
        Email = Regex.Replace(Email, "`(.*)`", "$1")

        If Email <> "" Then
            Email = SqlExpr(Email, """")
        End If
        Email = ExprExpr(" ", "", Email, "<" & SoEmail & ">")
        Return Email
    End Function

    Public Function SoEmailStr(ByVal Email As String) As String
        Return RegexGroup(Email, "(^|[ \t\[\<\>\""]*)([\w-.]+@[\w-]+(\.[\w-]+)+)(($|[ \t\<\>\""]*))", 2).Value
    End Function

    Function RegexGroup(ByVal Texto As String, ByVal Mascara As String, Optional ByVal Grupo As Object = 0) As System.Text.RegularExpressions.Group
        Return System.Text.RegularExpressions.Regex.Match(NZ(Texto, ""), Mascara).Groups(Grupo)
    End Function

    Sub Notifica(Optional ByVal Texto As String = "")
        System.Console.WriteLine("> " & Format(Now, "yyyy-MM-dd HH:mm") & " - " & IIf(Texto <> "", "", Replica.Status))
    End Sub

    Public Function EnviaEmail(ByVal De As String, ByVal Para As Object, ByVal Assunto As String, ByVal Corpo As String, Optional ByVal Prioridade As System.Net.Mail.MailPriority = MailPriority.Normal, Optional ByVal SmtpHost As String = Nothing, Optional ByVal SmtpPort As Integer = 25, Optional ByVal CC As Object = Nothing, Optional ByVal BCC As Object = Nothing, Optional ByVal SMTPUsuario As String = "", Optional ByVal SMTPSenha As String = "", Optional ByVal IncorporaImagens As Boolean = False, Optional ByVal UrlsLocais As ArrayList = Nothing, Optional ByVal Attachs As ArrayList = Nothing) As String
        Dim Mail As New MailMessage
        Dim Enviar As New System.Net.Mail.SmtpClient(NZ(SmtpHost, WebConf("smtp_host")), NZ(SmtpPort, WebConf("smtp_port")))
        Dim TMPS As New ArrayList
        Dim Ret As String = EnviaEmail(Mail, Enviar, De, Para, Assunto, Corpo, Prioridade, SmtpHost, SmtpPort, CC, BCC, SMTPUsuario, SMTPSenha, IncorporaImagens, , TMPS, UrlsLocais, Attachs)
        Mail.Dispose() ' libera arquivos
        ApagaTemps(TMPS)
        Return Ret
    End Function

    Sub ApagaTemps(ByVal Tmps As ArrayList)
        If Not IsNothing(Tmps) Then
            For Each tmp As String In Tmps
                Try
                    System.IO.File.Delete(tmp)
                Catch
                End Try
            Next
        End If
    End Sub

    Public Sub GravaLog(ByVal ArqLog As String, ByVal Msg As String, Optional ByVal IniciarArq As Boolean = False)
        For n As Integer = 1 To 10
            Try
                Using log As New System.IO.StreamWriter(ArqLog, Not IniciarArq)
                    log.WriteLine(Msg)
                    log.Close()
                End Using
                Exit Sub
            Catch
            End Try
            System.Threading.Thread.Sleep(10)
        Next
        Throw New Exception("Falha ao tentar gravar em arquivo de log uma ocorrência.")
    End Sub

    Function ReplRepl(ByVal Texto As String, ByVal De As String, ByVal Para As String) As String
        Do While InStr(Texto, De) <> 0
            Texto = Replace(Texto, De, Para)
        Loop
        Return Texto
    End Function

    Public Function TrimCarac(ByVal Texto As String, ByVal Carac() As String) As String
        Dim Achou As Boolean = True
        Do While Achou
            Achou = False
            For Each Item As String In Carac
                Do While Texto.StartsWith(Item, StringComparison.OrdinalIgnoreCase)
                    Texto = Mid(Texto, Len(Item) + 1)
                    Achou = True
                Loop
                Do While Texto.EndsWith(Item, StringComparison.OrdinalIgnoreCase)
                    Texto = StrStr(Texto, 0, -Len(Item))
                Loop
            Next
        Loop
        Return Texto
    End Function

    Function StrStr(ByVal Variavel As String, ByVal Inicio As Integer, Optional ByVal Final As Integer = Nothing) As String
        If Inicio < 0 Then
            Inicio = (Len(Variavel) + Inicio)
        End If
        If Not NZ(Final, 0) = 0 Then
            If Final < 0 Then
                Final = (Len(Variavel) + Final) - 1
            End If
            Return Variavel.Substring(Inicio, Final - Inicio + 1)
        End If
        Return Variavel.Substring(Inicio)
    End Function

    Public Function EnviaEmail(ByRef Mail As MailMessage, ByVal Enviar As System.Net.Mail.SmtpClient, ByVal De As String, ByVal Para As Object, ByVal Assunto As String, ByVal Corpo As String, ByVal ReplyTo As String, Optional ByVal Prioridade As System.Net.Mail.MailPriority = Nothing, Optional ByVal SmtpHost As String = Nothing, Optional ByVal SmtpPort As Integer = 25, Optional ByVal CC As Object = Nothing, Optional ByVal BCC As Object = Nothing, Optional ByVal SMTPUsuario As String = Nothing, Optional ByVal SMTPSenha As String = Nothing, Optional ByVal IncorporaImagens As Boolean = False, Optional ByRef CIDS As ArrayList = Nothing, Optional ByRef TMPS As ArrayList = Nothing, Optional ByVal UrlsLocais As ArrayList = Nothing, Optional ByVal Attachs As ArrayList = Nothing) As String
        Try


            ' cada param só é definido caso esteja mencionado
            If IsNothing(Mail) Then
                Mail = New MailMessage
            End If

            If Not IsNothing(De) Then
                Dim DeLista As ArrayList = TermosStrToLista(De)
                Mail.From = New MailAddress(EmailStr(DeLista(0)))
            End If

            If Not IsNothing(ReplyTo) Then
                Dim ReplyToLista As ArrayList = TermosStrToLista(ReplyTo)
                If ReplyToLista.Count > 0 Then
                    Mail.ReplyTo = New MailAddress(EmailStr(ReplyToLista(0)))
                End If
            End If

            If Not IsNothing(Para) Or Not IsNothing(CC) Or Not IsNothing(BCC) Then
                Mail.Bcc.Clear()
                Mail.CC.Clear()
                Mail.To.Clear()
            End If

            If Not IsNothing(Para) Then
                Dim ParaLista As ArrayList = TermosStrToLista(Para)
                For Each ParaItem As String In ParaLista
                    If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                        Dim M As New Email(ParaItem.Substring(4))
                        Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                    Else
                        Mail.To.Add(New MailAddress(EmailStr(ParaItem)))
                    End If
                Next
            End If

            If Not IsNothing(CC) Then
                Dim CCLista As ArrayList = TermosStrToLista(CC)
                For Each ParaItem As String In CCLista
                    If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                        Dim M As New Email(ParaItem.Substring(4))
                        Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                    Else
                        Mail.CC.Add(New MailAddress(EmailStr(ParaItem)))
                    End If
                Next
            End If

            If Not IsNothing(BCC) Then
                Dim BCCLista As ArrayList = TermosStrToLista(BCC)
                For Each ParaItem As String In BCCLista
                    If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                        Dim M As New Email(ParaItem.Substring(4))
                        Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                    Else
                        Dim M As New Email(ParaItem)
                        Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                    End If
                Next
            End If

            If Not IsNothing(Prioridade) Then
                Mail.Priority = Prioridade
            End If

            If Not IsNothing(Assunto) Then
                Mail.Subject = Assunto
            End If

            If Not IsNothing(Corpo) Then
                Mail.AlternateViews.Clear()


                If Not IncorporaImagens Then
                    Mail.IsBodyHtml = True
                    Mail.SubjectEncoding = System.Text.Encoding.GetEncoding("UTF-8")
                    Mail.BodyEncoding = System.Text.Encoding.GetEncoding("UTF-8")
                    Mail.Body = Corpo
                Else

                    ' inicia variávies de retorno caso não estejam definidas
                    If IsNothing(TMPS) Then
                        TMPS = New ArrayList
                    End If
                    If IsNothing(CIDS) Then
                        CIDS = New ArrayList
                    End If

                    ' visão alternativa
                    Dim alt As AlternateView = AlternateView.CreateAlternateViewFromString("", System.Text.Encoding.UTF8, "text/plain")
                    Mail.AlternateViews.Add(alt)

                    Dim arrImagens As New ArrayList
                    Dim listaImagens As String = "|"

                    For Each src As Match In Regex.Matches(Corpo, "url\(['|\""]+.*['|\""]\)|src=[""|'][^""']+[""|']", RegexOptions.IgnoreCase)
                        If InStr(1, listaImagens, "|" & src.Value & "|") = 0 Then
                            arrImagens.Add(src.Value)
                            listaImagens &= src.Value & "|"
                        End If
                    Next

                    CIDS.Clear()

                    For indx As Integer = 0 To arrImagens.Count - 1
                        Dim cid As String = "cid:EmbedRes_" & indx + 1
                        Corpo = Corpo.Replace(arrImagens(indx), "src=""" & cid & """")
                        Dim img As String = Regex.Replace(arrImagens(indx), "url\(['|\""]", "")
                        img = Regex.Replace(img, "src=['|\""]", "")
                        img = Regex.Replace(img, "['|\""]\)", "").Replace("""", "")


                        ' redirecionamentos
                        If Not IsNothing(UrlsLocais) Then
                            For Z = 0 To UrlsLocais.Count - 1 Step 2
                                Dim urlcomp As String = UrlsLocais(Z)
                                If img.StartsWith(urlcomp, StringComparison.OrdinalIgnoreCase) Then
                                    img = img.Replace(urlcomp, UrlsLocais(Z + 1))
                                End If
                            Next
                        End If

                        Dim URL As New System.Uri(img)
                        If URL.Scheme = "http" Or URL.Scheme = "ftp" Then
                            ' carrega imagens caso remotas

                            Try
                                Dim request As System.Net.HttpWebRequest = System.Net.WebRequest.Create(URL)

                                request.Timeout = 5000 ' cinco segundo de carga, senão erro...
                                Dim response As System.Net.HttpWebResponse = request.GetResponse()
                                Dim bmp As New Bitmap(response.GetResponseStream)

                                Dim DirTemp As String = NZV(WebConf("dir_temp"), "")
                                If DirTemp <> "" Then
                                    img = NomeArqLivre(DirTemp, "EnviaEmail")
                                Else
                                    img = System.IO.Path.GetTempFileName()
                                End If

                                If Not TMPS.Contains(img) Then
                                    TMPS.Add(img)
                                End If
                                bmp.Save(img)
                            Catch EX As Exception
                                Throw New Exception(EX.Message & " ao tentar obter conteúdo """ & URL.AbsolutePath & """")
                            End Try



                        End If
                        CIDS.Add(img)

                    Next

                    ' incorpora imagens
                    alt = AlternateView.CreateAlternateViewFromString(Corpo, System.Text.Encoding.UTF8, "text/html")

                    For z = 0 To CIDS.Count - 1
                        Dim res As New LinkedResource(CType(CIDS(z), String))
                        res.ContentId = "EmbedRes_" & z + 1
                        alt.LinkedResources.Add(res)
                    Next
                    Mail.AlternateViews.Add(alt)
                End If
            End If

            ' inclui attachados
            If Not IsNothing(Attachs) Then
                For Each attach As Object In Attachs
                    Try

                        If TypeOf attach Is String AndAlso attach <> "" Then
                            Mail.Attachments.Add(New Attachment(FileExpr(attach)))
                        ElseIf TypeOf attach Is System.Web.UI.WebControls.ListItem Then
                            Dim IT As System.Web.UI.WebControls.ListItem = attach
                            Dim ITA As New System.Net.Mail.Attachment(FileExpr(IT.Value))
                            ITA.Name = IT.Text
                            Mail.Attachments.Add(ITA)
                        End If
                    Catch
                    End Try
                Next
            End If

            If IsNothing(Enviar) Then
                Enviar = New System.Net.Mail.SmtpClient(NZ(SmtpHost, WebConf("smtp_host")), NZV(NZ(SmtpPort, WebConf("smtp_port")), 25))
            End If

            If Not IsNothing(SMTPUsuario) Then
                Enviar.Credentials = New System.Net.NetworkCredential(SMTPUsuario, NZ(SMTPSenha, ""))
            End If



            Enviar.Timeout = 100000
            Enviar.Send(Mail)
            Return ""
        Catch ex As Exception
            Return MessageEx(ex, "Erro ao tentar enviar email")
        End Try
    End Function

    Public Function TermosStrToLista(ByVal Email As Object) As ArrayList
        If TypeOf (Email) Is ArrayList Then
            Email = Join(CType(Email, ArrayList).ToArray, ";")
        End If
        Dim Lista As New ArrayList
        If NZ(Email, "") <> "" Then
            Dim Emails As Array = Split(Join(Split(Email, vbCrLf), ";"), ";")
            For Each Item As String In Emails
                Item = Trim(Item)
                If Item <> "" Then
                    Dim pref As String
                    If Item.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                        pref = "bcc:"
                        Item = Item.Substring(4)
                    Else
                        pref = ""
                    End If
                    If Item.StartsWith("conf.", StringComparison.OrdinalIgnoreCase) Then
                        Dim Result As ArrayList = TermosStrToLista(WebConf(Item.Substring(5)))
                        If pref <> "" Then
                            For Each ResultItem As String In Result
                                Lista.Add(pref & ResultItem)
                            Next
                        Else
                            Lista.AddRange(Result)
                        End If
                    Else
                        Lista.Add(pref & Item)
                    End If
                End If
            Next
        End If
        Return Lista
    End Function

    Public Class Email
        Private _completo As String = ""
        Private _soendereco As String = ""
        Private _descricao As String = ""
        Private _dominio As String = ""
        Private _primeironome As String = ""
        Private _ultimonome As String = ""

        ''' <summary>
        ''' Verifica a existência de caracteres inválidos para o formato de email padrão.
        ''' </summary>
        ''' <param name="Email">Email a ser verificado.</param>
        ''' <value>Endereço de email.</value>
        ''' <returns>True se email é válido ou false caso contrário.</returns>
        ''' <remarks></remarks>
        Shared ReadOnly Property Valida(ByVal Email As String) As Boolean
            Get
                Return Regex.IsMatch(EmailStr(Email), "(^|[ \t\[\<\>\""]*)([\w-.]+@[\w-]+(\.[\w-]+)+)(($|[ \t\<\>\""]*))")
            End Get
        End Property

        ''' <summary>
        ''' Verifica email carregado anteriormente.
        ''' </summary>
        ''' <value>True se email é válido ou false caso contrário.</value>
        ''' <returns>True se email é válido ou false caso contrário.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Valida() As Boolean
            Get
                Return Valida(_completo)
            End Get
        End Property

        ''' <summary>
        ''' Decompõe email em elementos.
        ''' </summary>
        ''' <param name="Email">Endereço de email.</param>
        ''' <remarks></remarks>
        Sub New(ByVal Email As String)
            _completo = EmailStr(Email)
            _soendereco = SoEmailStr(_completo)
            _descricao = Trim(RegexGroup(_completo, "\""(.*)\""", 1).Value)
            _dominio = RegexGroup(_soendereco, "@(.*)$", 1).Value

            Dim ems() As String = Split(_descricao & " ", " ")
            _primeironome = Trim(ems(0))
            _ultimonome = Trim(ems(ems.Length - 2))
        End Sub

        ''' <summary>
        ''' Domínio do email.
        ''' </summary>
        ''' <value>Domínio do email (depois do arroba).</value>
        ''' <returns>Domínio do email.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Dominio() As String
            Get
                Return _dominio
            End Get
        End Property

        ''' <summary>
        ''' Email completo já formatado.
        ''' </summary>
        ''' <value>Email completo já formatado.</value>
        ''' <returns>Email completo já formatado.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Completo() As String
            Get
                Return _completo
            End Get
        End Property

        ''' <summary>
        ''' Só o endereço do email (antes do arroba).
        ''' </summary>
        ''' <value>Só o endereço do email.</value>
        ''' <returns>Só o endereço do email.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SoEndereco() As String
            Get
                Return _soendereco
            End Get
        End Property

        ''' <summary>
        ''' Descrição do email (trecho entre apóstrofos antes do email).
        ''' </summary>
        ''' <value>Descrição do email (trecho entre apóstrofos antes do email).</value>
        ''' <returns>Descrição do email (trecho entre apóstrofos antes do email).</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Descricao() As String
            Get
                Return _descricao
            End Get
        End Property

        ''' <summary>
        ''' Primeiro nome na descrição do email.
        ''' </summary>
        ''' <value>Primeiro nome na descrição do email.</value>
        ''' <returns>Primeiro nome na descrição do email.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PrimeiroNome() As String
            Get
                Return _primeironome
            End Get
        End Property

        ''' <summary>
        ''' Último nome na descrição de email.
        ''' </summary>
        ''' <value>Último nome na descrição de email.</value>
        ''' <returns>Último nome na descrição de email.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property UltimoNome() As String
            Get
                Return _ultimonome
            End Get
        End Property
    End Class


    Function NomeArqLivre(ByVal NomeDir As String, ByVal NomeArq As String) As String
        Dim DD As New System.IO.DirectoryInfo(NomeDir)
        If Not DD.Exists Then
            DD.Create()
            NomeArq = FileExpr(NomeDir, NomeArq)
        Else
            Dim z As Integer = 1
            Dim NomeTest As String = FileExpr(NomeDir, NomeArq)
            Do While True
                If z <> 1 Then
                    NomeTest = FileExpr(NomeDir, System.IO.Path.GetFileNameWithoutExtension(NomeArq) & "_" & Trim(Format(z, "    00")) & System.IO.Path.GetExtension(NomeArq))
                End If
                Dim FF As New System.IO.FileInfo(NomeTest)
                If Not FF.Exists Then
                    Exit Do
                End If
                z += 1
            Loop
            NomeArq = NomeTest
        End If
        Return NomeArq
    End Function

    Public Function MessageEx(ByVal Ex As Exception, Optional ByVal MensagemCompl As String = "") As String

        ' mensagem padrão
        Dim Mensagem As String = Ex.Message

        If Not IsNothing(Ex.InnerException) AndAlso NZ(Ex.InnerException.Message, "") <> "" Then
            Mensagem &= ". " & Ex.InnerException.Message
        End If
        Dim Param As String

        ' mensagens específicas
        Param = RegexGroup(Mensagem, "Cannot update (.*); field not updateable", 1).Value
        If Param <> "" Then
            Mensagem = "Por restrições da base de dados, campo " & Param & " não pode ser atualizado"
        End If

        Param = RegexGroup(Mensagem, "create duplicate values in the").Value
        If Param <> "" Then
            Mensagem = "Tentativa de registro de chave duplicada"
        End If

        Param = RegexGroup(Mensagem, "Cannot set column (.*). The value violates the MaxLength.*", 1).Value
        If Param <> "" Then
            Mensagem = "Tamanho do campo " & Param & " excede o limite"
        End If

        Param = RegexGroup(Mensagem, "The path is not of a legal").Value
        If Param <> "" Then
            Mensagem = "Caminho de arquivo inexistente ou ilegal"
        End If

        Param = RegexGroup(Mensagem, "Duplicate entry (.*) for key .*", 1).Value
        If Param <> "" Then
            Mensagem = "Tentativa de gravação de registro duplicado - " & Param
        End If

        Param = RegexGroup(Mensagem, "Empty path name is not legal").Value
        If Param <> "" Then
            Mensagem = "Nome de arquivo incorreto"
        End If

        Param = RegexGroup(Mensagem, "Could not find file '(.*?)'", 1).Value
        If Param <> "" Then
            Mensagem = "Arquivo não encontrado: " & Param
        End If

        Param = RegexGroup(Mensagem, "Thread was being aborted|O thread estava sendo anulado").Value
        If Param <> "" Then
            Mensagem = "É necessário logar-se ou sua sessão foi encerrada."
        End If


        ' ------------------------------------------------
        ' TRATAMENTO DE ERROS DO ORACLE
        If InStr(Mensagem, "ORA-01400:") <> 0 Then
            Mensagem = "Campo de identificação do registro não pode estar nulo"
        End If

        Param = RegexGroup(Mensagem, "ORA-00372:").Value
        If Param <> "" Then
            Mensagem = "Base de dados em condição de apenas para leitura ou parada para manutenção. Informe sua necessidade ao suporte"
        End If

        Param = RegexGroup(Mensagem, "ORA-02291: .*\((.*)\)", 1).Value
        If Param <> "" Then
            Mensagem = "Falta de registro relacionado em " & Param
        End If

        Param = RegexGroup(Mensagem, "ORA-00001: .*\((.*)\)", 1).Value
        If Param <> "" Then
            Mensagem = "Tentativa de registro de chave duplicada em " & Param
        End If

        Param = RegexGroup(Mensagem, "ORA-01017:").Value
        If Param <> "" Then
            Mensagem = "Logon incorreto. Usuário ou senha inválidos ou sessão expirada"
        End If

        Param = RegexGroup(Mensagem, "ORA-00942:").Value
        If Param <> "" Then
            Mensagem = "Tabela ou visão inexistente"
        End If

        Param = RegexGroup(Mensagem, "ORA-12541:|ORA-12170:").Value
        If Param <> "" Then
            Mensagem = "Banco de dados indisponível no momento. Suporte já foi contactado"
        End If

        ' ------------------------------------------------
        ' TRATAMENTO DE ERROS DO MYSQL

        Param = RegexGroup(Mensagem, "Access denied for user (.*)", 1).Value
        If Param <> "" Then
            Mensagem = "Acesso não autorizado para " & Param & ". Verifique usuário e senha e tente novamente"
        End If

        Mensagem = IIf(MensagemCompl <> "", MensagemCompl & ". ", "") & Mensagem & "."
        Return Mensagem
    End Function

    Public Function VersaoApl() As String
        Dim v() As String = Split(My.Application.Info.Version.ToString & ".0.0.0.0", ".")
        Return "V" & Format(Val(v(0)), "00") & "." & Format(Val(v(1)), "00") & "." & Format(Val(v(2)), "00") & "." & Format(Val(v(3)), "00")
    End Function

End Module
