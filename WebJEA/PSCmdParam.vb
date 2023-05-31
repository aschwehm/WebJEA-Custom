Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports Amazon.Runtime
Imports System.CodeDom.Compiler

Public Class PSCmdParam
    Private dlog As NLog.Logger = NLog.LogManager.GetCurrentClassLogger()

    Public Enum ParameterType
        PSString
        PSInt
        PSFloat
        PSDate
        PSBoolean
        PSButton
    End Enum

    Public Name As String
    Public VisibleName As String = ""
    Public PostBackVisibleName As String = ""
    Public HelpMessage As String = ""
    Public FormGroup As String = ""
    Public HelpDetail As String = ""
    Public AutoPostBack As Boolean = False
    Public BackLinkFormGroup As String = ""
    Public DirectiveMultiline As Boolean = False
    Public DirectiveDateTime As Boolean = False
    Public VarType As String = ""
    Public DefaultValue As Object = Nothing
    Public DefaultValue_BT As Object = Nothing
    'TODO: Add support for more than just string default values - can we support arrays?
    Public Validation As New List(Of String)
    'Private prvValidation As New List(Of PSCmdParamVal)

    Sub New()

    End Sub

    Public ReadOnly Property IsMandatory As Boolean
        Get
            For Each val As String In Validation
                If val.ToUpper = "MANDATORY" Then
                    'parameter is required
                    Return True
                End If
            Next
            Return False
        End Get
    End Property

    Public ReadOnly Property ParamType As ParameterType
        Get
            Dim vartypestr = VarType.ToLower
            If vartypestr Like "string*" Then
                Return ParameterType.PSString
            ElseIf vartypestr = "datetime" Then
                Return ParameterType.PSDate
            ElseIf vartypestr Like "single*" Or vartypestr Like "double*" Or vartypestr Like "float*" Then
                Return ParameterType.PSFloat
            ElseIf vartypestr Like "bool*" Or vartypestr Like "switch" Then
                Return ParameterType.PSBoolean
            ElseIf vartypestr Like "int*" Or vartypestr Like "uint*" Or vartypestr Like "byte*" Or vartypestr Like "long*" Then
                Return ParameterType.PSInt
            ElseIf vartypestr = "button" Then
                Return ParameterType.PSButton
            End If
            'by default we treat a value as string.  This includes PSCredential
            Return ParameterType.PSString
        End Get
    End Property

    Public ReadOnly Property IsMultiValued As Boolean
        Get
            If VarType.Contains("[]") Then
                Return True
            ElseIf DirectiveMultiline Then
                Return False
            End If
            Return False
        End Get
    End Property

    Public ReadOnly Property AllowedValues As List(Of String)
        Get
            'this param does not explicitly define allowed values
            If IsSelect = False Then Return Nothing
            For Each valobj As PSCmdParamVal In ValidationObjects
                If valobj.Type = PSCmdParamVal.ValType.SetCol Then
                    Return valobj.Options
                End If
            Next

            'should not have gotten here
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property IsSelect As Boolean
        Get
            For Each valobj As PSCmdParamVal In ValidationObjects
                If valobj.Type = PSCmdParamVal.ValType.SetCol Then
                    'parameter is restricted
                    Return True
                End If
            Next
            'no validateset found
            Return False
        End Get
    End Property

    Public ReadOnly Property ValidationObjects As List(Of PSCmdParamVal)
        Get
            Dim retobjs As New List(Of PSCmdParamVal)
            For Each rule As String In Validation
                Dim obj As New PSCmdParamVal(rule)
                If obj.IsValid Then
                    retobjs.Add(obj)
                End If
            Next
            Return retobjs

        End Get
    End Property

    Public Sub AddValidation(valstring As String)
        'TODO add support for properly managing conflicting validation options
        If valstring.ToUpper.StartsWith("VALIDATE") Or valstring.ToUpper.StartsWith("ALLOW") Or valstring.ToUpper.StartsWith("MANDATORY") Then
            If Not Validation.Contains(valstring) Then
                'don't add precise duplicates.  Doesn't stop from adding incompatible validation commands
                Validation.Add(valstring)
            End If
        ElseIf valstring.ToUpper.StartsWith("ALIAS") Then
            'do nothing
        Else 'variable
            dlog.Warn("Unexpected Validation Type not supported: " & valstring)
        End If

    End Sub

    Public Function Clone() As PSCmdParam

        Dim psparam As New PSCmdParam
        'Dim CachedValue As String
        psparam.Name = Name
        psparam.BackLinkFormGroup = BackLinkFormGroup
        psparam.AutoPostBack = AutoPostBack

        If FormGroup IsNot "" Then
            psparam.FormGroup = FormGroup
        End If

        '---------
        Dim SessionID = HttpContext.Current.Session.SessionID

        If (WebJEA._default.SessionValues.Item(SessionID).ContainsKey("UPDATE_psparam_" + FormGroup)) Then
            WebJEA._default.SessionValues.Item(SessionID).Remove("UPDATE_psparam_" + FormGroup)
            If AutoPostBack = False Then
                psparam.AutoPostBack = True

                If WebJEA._default.SessionValues.Item(SessionID).ContainsKey("REFRESH_psparam_" + Name) Then
                    WebJEA._default.SessionValues.Item(SessionID).Remove("REFRESH_psparam_" + Name)
                    WebJEA._default.SessionValues.Item(SessionID).Add("REFRESH_psparam_" + Name, "")
                End If
            End If
            If WebJEA._default.SessionValues.Item(SessionID).ContainsKey("psparam_" + Name) Then
                WebJEA._default.SessionValues.Item(SessionID).Remove("psparam_" + Name)
            End If
            If WebJEA._default.SessionValues.Item(SessionID).ContainsKey("ALLVARS_" + Name) Then
                WebJEA._default.SessionValues.Item(SessionID).Remove("ALLVARS_" + Name)
            End If
        End If
        '---------

        'Set the PostBack Value as Default Value in the Grouped FormField
        If FormGroup IsNot "" Then
            If WebJEA._default.SessionValues.Item(SessionID) IsNot Nothing Then
                'If (WebJEA._default.SessionValues.Item(SessionID).Item("psparam_" + BackLinkFormGroup) Is Nothing And Not BackLinkFormGroup = "") Then
                If Not (WebJEA._default.SessionValues.Item(SessionID).ContainsKey("psparam_" + FormGroup)) Then
                    DefaultValue = Nothing
                Else
                    DefaultValue = WebJEA._default.SessionValues.Item(SessionID).Item("psparam_" + FormGroup)
                End If
                'psparam.FormGroup = FormGroup
            End If
        ElseIf (WebJEA._default.SessionValues.Item(SessionID).ContainsKey("psparam_" + Name)) Then
            DefaultValue = WebJEA._default.SessionValues.Item(SessionID).Item("psparam_" + Name)
        Else
            DefaultValue = DefaultValue

        End If


        If VisibleName = "" Then
            psparam.VisibleName = Name
        Else
            psparam.VisibleName = VisibleName
        End If


        psparam.HelpMessage = HelpMessage
        psparam.HelpDetail = HelpDetail
        psparam.VarType = VarType
        psparam.DirectiveDateTime = DirectiveDateTime
        psparam.DirectiveMultiline = DirectiveMultiline

        psparam.PostBackVisibleName = PostBackVisibleName

        'Sobald in der validate.ps1 eine Variable mit FPIT am Anfang gefunden wird dann spring er aus seinem ursprünglichen Script

        If psparam.Name.Substring(0, 4) = "FPIT" And Not WebJEA._default.SessionValues.Item(SessionID).Contains("ALLVARS_" + Name) Then

            Dim Ausgabe As String = ""

            If (FormGroup Is "" And DefaultValue Is Nothing) Or (FormGroup IsNot "" And DefaultValue IsNot Nothing) Then

                If WebJEA._default.SessionValues.Item(SessionID).Contains("EXEC_" + Name) Then 'If its a Button and has a DefaultValue from a FormGroup Field then execute the additional Script
                    WebJEA._default.SessionValues.Item(SessionID).Remove("EXEC_" + Name)

                    Dim FPIT_Path = WebJEA.My.Settings.configfile
                    FPIT_Path = FPIT_Path.Replace("config.json", Name + ".ps1")
                    Dim pscommand As String = FPIT_Path & " " & DefaultValue & "; exit $LASTEXITCODE"
                    Dim cmd As String = "powershell.exe -noprofile -NonInteractive -WindowStyle hidden -command " & pscommand
                    Dim shell = CreateObject("WScript.Shell")
                    Dim executor = shell.Exec(cmd)
                    executor.StdIn.Close

                    psparam.DefaultValue_BT = executor.StdOut.ReadAll


                Else 'If its not a button then normally execute the external Script to prefill Fields
                    If DefaultValue = "" Or DefaultValue = Nothing And Not (ParamType = ParameterType.PSButton) Then
                        DefaultValue = "0"
                    End If

                    'Aufruf externes Powershell Script
                    'script.ps1 <VariablenName> <Wert>  C:\temp\WebJEA\WebJea-Scripts\test.ps1
                    Dim FPIT_Path = WebJEA.My.Settings.configfile
                    FPIT_Path = FPIT_Path.Replace("config.json", "FPIT_Commands.ps1")
                    Dim pscommand As String = FPIT_Path & " " & psparam.Name & " '" & DefaultValue & "'" & "; Exit $LASTEXITCODE"
                    Dim cmd As String = "powershell.exe -noprofile -NonInteractive -WindowStyle hidden -command " & pscommand
                    Dim shell = CreateObject("WScript.Shell")
                    Dim executor = shell.Exec(cmd)
                    executor.StdIn.Close

                    Ausgabe = executor.StdOut.ReadAll

                    'Filter Special Characters
                    Ausgabe = Ausgabe.Replace("Ã¤", "ä")
                    Ausgabe = Ausgabe.Replace("Ã¼", "ü")
                    Ausgabe = Ausgabe.Replace("Ã¶", "ö")

                    Ausgabe = Ausgabe.Replace("Ã„", "Ä")
                    Ausgabe = Ausgabe.Replace("Ãœ", "Ü")
                    Ausgabe = Ausgabe.Replace("Ã–", "Ö")

                    Ausgabe = Ausgabe.Replace("ÃŸ", "ß")
                End If
            Else
                Ausgabe = Nothing
            End If

            If psparam.Name.Substring(4, 2) = "SL" Then

                'SingleLine (Textbox), einfache Ausgabe
                psparam.DefaultValue = Ausgabe

            ElseIf psparam.Name.Substring(4, 2) = "ML" Then

                'Multiline
                psparam.DirectiveMultiline = True

                psparam.DefaultValue = Ausgabe
            ElseIf psparam.Name.Substring(4, 2) = "BT" Then

                'Button
                psparam.DefaultValue = DefaultValue

            ElseIf (psparam.Name.Substring(4, 2) = "LB") Or (psparam.Name.Substring(4, 2) = "LS") Then

                Dim arr = Split(Ausgabe, ";")
                ReDim Preserve arr(UBound(arr) - 1)

                Dim tmpValidation As String = "VALIDATESET("
                Dim MyInList As New System.Collections.Generic.List(Of String)

                For Each Eintrag As String In arr
                    MyInList.Add(Eintrag)
                    tmpValidation = tmpValidation + "'" + Eintrag + "',"
                Next

                tmpValidation = Left(tmpValidation, (tmpValidation.Length - 1))
                tmpValidation = tmpValidation + ")"

                If tmpValidation = "VALIDATESET)" Then
                    tmpValidation = "VALIDATESET('Keine Daten')"
                    psparam.AddValidation(tmpValidation)
                Else
                    psparam.AddValidation(tmpValidation)
                    WebJEA._default.SessionValues.Item(SessionID).Add("ALLVARS_" + Name, tmpValidation)
                End If


            Else

            End If

            '###########

        ElseIf psparam.Name.Substring(0, 4) = "FPIT" And WebJEA._default.SessionValues.Item(SessionID).Contains("ALLVARS_" + Name) Then
            psparam.AddValidation(WebJEA._default.SessionValues.Item(SessionID).Item("ALLVARS_" + Name))
            If WebJEA._default.SessionValues.Item(SessionID).Contains("psparam_" + Name) Then
                psparam.DefaultValue = WebJEA._default.SessionValues.Item(SessionID).Item("psparam_" + Name)
            Else
                psparam.DefaultValue = Nothing
            End If
        Else 'Rest der alten Funktion
            psparam.DefaultValue = DefaultValue

            For Each val As String In Validation
                psparam.AddValidation(val)
            Next
        End If

        Return psparam

    End Function


    Public Sub MergeUnder(psparam As PSCmdParam)
        'this will merge "under" the current parameter.
        'it will NOT overwrite properties (Help, etc), but if there is no value specified, it will add the value

        If String.IsNullOrWhiteSpace(HelpMessage) Then HelpMessage = psparam.HelpMessage
        If String.IsNullOrWhiteSpace(HelpDetail) Then HelpDetail = psparam.HelpDetail
        If String.IsNullOrWhiteSpace(VarType) Then VarType = psparam.VarType
        If String.IsNullOrWhiteSpace(DefaultValue) Then DefaultValue = psparam.DefaultValue
        For Each valstr As String In psparam.Validation
            AddValidation(valstr)
        Next

    End Sub
    Public Sub MergeOver(psparam As PSCmdParam)
        'this will merge "over" the current parameter.
        'it WILL overwrite properties (Help, etc), if specified
        'validation is always merge

        If Not String.IsNullOrWhiteSpace(psparam.HelpMessage) Then HelpMessage = psparam.HelpMessage
        If Not String.IsNullOrWhiteSpace(psparam.HelpDetail) Then HelpDetail = psparam.HelpDetail
        If Not String.IsNullOrWhiteSpace(psparam.VarType) Then VarType = psparam.VarType
        If Not String.IsNullOrWhiteSpace(psparam.DefaultValue) Then DefaultValue = psparam.DefaultValue
        For Each valstr As String In psparam.Validation
            AddValidation(valstr)
        Next

    End Sub

    Public ReadOnly Property FieldName As String
        Get
            Return "psparam_" & Name
        End Get
    End Property

End Class
