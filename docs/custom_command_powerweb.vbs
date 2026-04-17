Option Explicit

' =====================================================================
' Comando: PruebaComando
' Entorno: Meridian Web / PowerWeb
' Objetivo:
'   - Pedir al usuario la especialidad usando input web-compatible (Batch)
'   - Validar el dato en PreExecute
'   - Escribir el valor en Document.Especialidad en Execute
'
' Nota:
'   Este script NO usa WinMsgBox/InputBox.
'   Ajusta únicamente los métodos de Batch marcados en Adapter_* si en tu
'   versión de Meridian tienen un nombre distinto.
' =====================================================================

Private Const PARAM_ESPECIALIDAD As String = "especialidad"
Private Const LABEL_ESPECIALIDAD As String = "Especialidad"
Private Const DEFAULT_ESPECIALIDAD As String = "CIVIL"
Private Const MAX_LEN As Integer = 30

Private mEspecialidad

' ---------------------------------------------------------------------
' PreInitialize
' Define el campo web que se mostrará al usuario antes de ejecutar.
' ---------------------------------------------------------------------
Sub PruebaComando_PreInitialize(Batch)
    On Error GoTo EH

    Call Adapter_RegisterTextInput(Batch, PARAM_ESPECIALIDAD, LABEL_ESPECIALIDAD, DEFAULT_ESPECIALIDAD, True, MAX_LEN)

    Exit Sub
EH:
    Call Adapter_Fail(Batch, "PreInitialize: " & Err.Description)
End Sub

' ---------------------------------------------------------------------
' PreExecute
' Lee y valida el valor ingresado por el usuario.
' ---------------------------------------------------------------------
Sub PruebaComando_PreExecute(Batch)
    On Error GoTo EH

    mEspecialidad = UCase(Trim(CStr(Adapter_ReadInputValue(Batch, PARAM_ESPECIALIDAD, DEFAULT_ESPECIALIDAD))))

    If Len(mEspecialidad) = 0 Then
        Call Adapter_Invalidate(Batch, "Debe ingresar una especialidad.")
        Exit Sub
    End If

    If Len(mEspecialidad) > MAX_LEN Then
        Call Adapter_Invalidate(Batch, "La especialidad supera " & CStr(MAX_LEN) & " caracteres.")
        Exit Sub
    End If

    Exit Sub
EH:
    Call Adapter_Fail(Batch, "PreExecute: " & Err.Description)
End Sub

' ---------------------------------------------------------------------
' Execute
' Aplica el valor validado sobre el objeto destino.
' ---------------------------------------------------------------------
Sub PruebaComando_Execute(Batch)
    On Error GoTo EH

    ' Seguridad adicional por si el runtime no mantiene estado de módulo.
    If IsEmpty(mEspecialidad) Or Len(Trim(CStr(mEspecialidad))) = 0 Then
        mEspecialidad = UCase(Trim(CStr(Adapter_ReadInputValue(Batch, PARAM_ESPECIALIDAD, DEFAULT_ESPECIALIDAD))))
    End If

    If Not Document Is Nothing Then
        Document.Especialidad = mEspecialidad

    ElseIf Not Folder Is Nothing Then
        ' Si el comando se lanza sobre carpeta, deja trazabilidad y no falla.
        Call Adapter_AddInfo(Batch, "El comando fue invocado sobre carpeta; no se actualiza Especialidad.")

    Else
        Call Adapter_Invalidate(Batch, "No hay objeto Document ni Folder en contexto.")
    End If

    Exit Sub
EH:
    Call Adapter_Fail(Batch, "Execute: " & Err.Description)
End Sub

' ========================
' Adaptadores Batch/API
' ========================

Private Sub Adapter_RegisterTextInput(Batch, key, label, defaultValue, required, maxLen)
    On Error Resume Next

    ' Variante A: parámetros en colección Batch.Parameters
    Batch.Parameters(key).Label = label
    Batch.Parameters(key).DefaultValue = defaultValue
    Batch.Parameters(key).Required = required
    Batch.Parameters(key).MaxLength = maxLen
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    ' Variante B: método declarativo
    Batch.AddInput key, label, "string", required, maxLen, defaultValue
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    ' Variante C: método simplificado
    Batch.AddParameter key, defaultValue
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    On Error GoTo 0
    Err.Raise vbObjectError + 1100, "Adapter_RegisterTextInput", "No se pudo registrar el input web-compatible."
End Sub

Private Function Adapter_ReadInputValue(Batch, key, defaultValue)
    On Error Resume Next

    Adapter_ReadInputValue = Batch.Parameters(key).Value
    If Err.Number = 0 Then
        If Len(Trim(CStr(Adapter_ReadInputValue))) = 0 Then Adapter_ReadInputValue = defaultValue
        Exit Function
    End If
    Err.Clear

    Adapter_ReadInputValue = Batch.GetValue(key)
    If Err.Number = 0 Then
        If Len(Trim(CStr(Adapter_ReadInputValue))) = 0 Then Adapter_ReadInputValue = defaultValue
        Exit Function
    End If
    Err.Clear

    Adapter_ReadInputValue = defaultValue
End Function

Private Sub Adapter_Invalidate(Batch, message)
    On Error Resume Next

    Batch.IsValid = False
    Batch.Cancel = True
    Batch.ErrorMessage = message

    If Err.Number <> 0 Then
        Err.Clear
        Batch.AddError message
    End If
End Sub

Private Sub Adapter_Fail(Batch, message)
    On Error Resume Next
    Call Adapter_Invalidate(Batch, message)
End Sub

Private Sub Adapter_AddInfo(Batch, message)
    On Error Resume Next

    Batch.AddMessage message
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub
