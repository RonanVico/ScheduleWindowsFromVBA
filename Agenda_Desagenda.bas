Attribute VB_Name = "Agenda_Desagenda"
Option Compare Database
Option Explicit



Public Sub teste()
    Call AgendaTarefaWindows("DownloadExtratosVBA", "C:\Teste\Teste.txt", DateTime.Date, DateTime.time)
End Sub




Public Function AgendaTarefaWindows(ByVal taskname$, ByVal ArquivoQueVaiRodar$, ByVal data$, ByVal HORARIO$, Optional ByVal sc$ = "ONCE") As Boolean
'FEITO POR RONAN VICO EL MAGO - https://www.linkedin.com/in/ronan-vico/


'/ts = taskname
'/tr = Task Run (o que vai rodar
'sc = Schedule Frequency , frequencia que vai checar
'/sc  = MINUTE, HOURLY, DAILY, WEEKLY, MONTHLY, ONCE, ONSTART, ONLOGON, ONIDLE.
'/sd = Schedule DATE , FORMATO ("DD/mm/AAAA")
'/st = Schedule TIME , formato 00:00:00
Dim localArquivoQueVaiRodar As String
localArquivoQueVaiRodar = ArquivoQueVaiRodar
On Error GoTo err_handler
    Dim WSShell As Object
    Dim WSShellExec As Object
    Dim hwnd
    Dim out
    Dim comando As String
    Dim accessPath As String
    HORARIO = Replace(HORARIO, ":", "")
    HORARIO = Mid(HORARIO, 1, 2) & ":" & Mid(HORARIO, 3, 2) & ":" & Mid(HORARIO, 5, 2)
    
    'DesAgendaTarefaWindows (taskname)
    'ArquivoQueVaiRodar = "c:\teste\teste.txt"
    
    Dim fso As Object
    Dim fsFolder As Object
    Dim fsFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsFile = fso.GetFile(localArquivoQueVaiRodar)
    localArquivoQueVaiRodar = fsFile.shortpath
    accessPath = (SysCmd(acSysCmdAccessDir)) & "msaccess.exe"
    Set fsFile = fso.GetFile(accessPath)
    'accessPath = fsFile.shortpath
    'If Right(ArquivoQueVaiRodar, "3") = "ACC" Then ArquivoQueVaiRodar = Replace(ArquivoQueVaiRodar, ".acc", ".accdb")
    comando = Replace("/create /tn |" & taskname & "| /tr |'" & accessPath _
              & "' '" & localArquivoQueVaiRodar & "'|   /sc |" & sc & "| /sd |" & data & "|  /st |" & HORARIO & "|  /f", "|", Chr(34))
              
    comando = Replace("/create /tn |" & taskname & "| /tr |\|" & accessPath _
              & "\| " & localArquivoQueVaiRodar & "|   /sc |" & sc & "| /sd |" & data & "|  /st |" & HORARIO & "|  /f", "|", Chr(34))
          
    'ShellExecute 0, "runas", "schtasks", comando, vbNullString, 1
    ShellExecute 0, "runas", "schtasks", comando, vbNullString, 1
    
    
    Set WSShell = CreateObject("WScript.shell")
    Set WSShell = WSShell.exec("schtasks /query /tn " & Chr(34) & taskname & Chr(34))
    out = WSShell.stdout.readall
    
    'Se tiver dentro do access após ele ter sido executado pelo sch, o aplicativo fica em running.
    If InStr(out, "Em exec") <> 0 Or InStr(out, "running") <> 0 Then
        Call DesAgendaTarefaWindows(taskname)
        AgendaTarefaWindows = AgendaTarefaWindows(taskname, ArquivoQueVaiRodar, data, HORARIO, sc)
    Else
        If InStr(out, data) <> 0 And (InStr(out, "Pronto") <> 0 Or InStr(out, "Ready") <> 0) And InStr(out, HORARIO) Then
            MsgBox out, , "Resposta do windows!"
            AgendaTarefaWindows = True
        End If
    End If
    
    
    Set WSShell = Nothing
    Set fs = Nothing
    Set fso = Nothing
    Exit Function
    Resume
    
err_handler:
    MsgBox err.Number & "  " & err.Description
    
End Function

Public Function DesAgendaTarefaWindows(ByVal taskname As String) As Boolean
'FEITO POR RONAN VICO EL MAGO - https://www.linkedin.com/in/ronan-vico/

On Error GoTo err_handler
    Dim WSShell As Object
    Dim WShellExec As Object
    Dim hwnd
    Dim out
    Dim comando As String
    
              
    'ShellExecute 0, "runas", "schtasks", comando, vbNullString, 1
    
    
    comando = "/delete /tn " & taskname & " /f"
    ShellExecute 0, "runas", "schtasks", comando, vbNullString, 1
    Set WSShell = CreateObject("WScript.shell")
    Set WShellExec = WSShell.exec("schtasks /query /tn " & Chr(34) & taskname & Chr(34))
    
    
    out = WShellExec.stdout.readall
    
    If out = "" Then
        DesAgendaTarefaWindows = True
    End If
    Set WSShell = Nothing
    Set WShellExec = Nothing
    Exit Function
    Resume
    
err_handler:
    Stop

End Function


