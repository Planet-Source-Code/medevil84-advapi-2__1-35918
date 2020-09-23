Attribute VB_Name = "ModElab"
Enum typeOperation
    typeOpen = 1
    typeAdd = 0
End Enum

Sub FillLists(TextLoaded As String, Api As clsAPI, Optional bar As ProgressBar, Optional pStatusText As Panel)
    
    'Load elements...
    Dim vet() As String
    
    vet() = Split(TextLoaded, vbCrLf)
    
    'TypeTrue = False
    
    On Error Resume Next
    pStatusText.Text = lng_Loading
    On Error GoTo 0
        
    If IsMissing(bar) = True Then
        For i = LBound(vet) To UBound(vet)
            Call SelectType(vet(i), Api)
        Next i
    Else
        On Error Resume Next
        bar.Max = UBound(vet) + 1
        bar.Min = LBound(vet) + 1
        For i = LBound(vet) To UBound(vet)
            Call SelectType(vet(i), Api)
            bar.Value = i + 1
        Next i
        bar.Value = bar.Min
        On Error GoTo 0
    End If
    
End Sub

Public Sub SelectType(Text As String, Api As clsAPI)

    Dim spWord1 As String, spWord2 As String
    Static TypeTrue As Boolean
    
    If Text = "" Then Exit Sub
    If Left(Trim(Text), 1) = "'" Then Exit Sub
    
    spWord1 = LCase(PrimaParola(Text))
    spWord2 = RestoStringa(Text)

    Select Case spWord1
    Case "global"
        SelectType spWord2, Api
    Case "public"
        SelectType spWord2, Api
    Case "private"
        SelectType spWord2, Api
    
    Case "declare"
        Call AddDeclare(spWord2, Api)
    Case "type"
        TypeTrue = True
        Call AddType(spWord2, Api, True)
    Case "const"
        Call AddConst(spWord2, Api)
    Case Else
        If TypeTrue = True Then
            If spWord1 = "end" And LCase(spWord2) = "type" Then
                TypeTrue = False
            ElseIf Not Left(Trim(Text), 1) = "'" Then
                Call AddType(Text, Api, False)
            End If
        'Else: Assume it's a comment
        End If
    End Select
End Sub

Public Sub AddType(Text As String, Api As clsAPI, opNew As Boolean)

    Dim X As New apiType, Y As New apiParams
    
    If Not opNew Then
        'Get Object
        Set X = Api.cTypes(Api.cTypes.Count)
        
        'if text = "x as long"
        'PrimaParola(text) = "x"
        'RestoStringa(RestoStringa(text)) = "long"
        
        Y.paramName = PrimaParola(Text)
        Y.paramType = RestoStringa(RestoStringa(Text))
        
        'Add a new key
        Y.idKey = "par" & X.decParamsID
        
        'Incremento contatore
        X.decParamsID = X.decParamsID + 1
        
        X.decParams.Add Y, Y.idKey
        
        Set X = Nothing
        Set Y = Nothing
        
    Else
        
        X.decName = Trim(Text)
        
        'Add a new key
        X.idKey = "type" & Api.cTypesID
        Api.cTypesID = Api.cTypesID + 1
        
        Api.cTypes.Add X, X.idKey
        
        Set X = Nothing
    End If
End Sub
Public Sub AddConst(Text As String, Api As clsAPI)
    Dim X As New apiConst
    
    X.decName = PrimaParola(Text)
    If InStr(1, LCase(Text), " as ", vbTextCompare) > 0 Then
        X.decType = PrimaParola(RestoStringa(RestoStringa(Text)))
        X.decValue = PrimaParola(RestoStringa(RestoStringa(RestoStringa(RestoStringa(Text)))))
    Else
        X.decValue = PrimaParola(RestoStringa(RestoStringa(Text)))
    End If
    
    'Add a new key
    X.idKey = "const" & Api.cConstsID
    Api.cConstsID = Api.cConstsID + 1
    
    Api.cConsts.Add X, X.idKey
End Sub
Public Sub AddDeclare(Text As String, Api As clsAPI)
    Dim X As New apiDeclares, Y As New apiParams
    Dim f() As String, p1 As Long, p2 As Long, i As Long
    Dim r As String
    
    X.decSub = CBool(LCase(PrimaParola(Text)) = "sub")
    Text = RestoStringa(Text)
    X.decName = PrimaParola(Text)
    Text = RestoStringa(Text)
    X.decLib = LCase(Replace(PrimaParola(RestoStringa(Text)), Chr(34), ""))
    If Not CBool(InStr(1, X.decLib, ".dll", vbTextCompare) Or _
       InStr(1, X.decLib, ".drv", vbTextCompare)) Then
            X.decLib = X.decLib & ".dll"
    End If
        
    Text = RestoStringa(RestoStringa(Text))
    
    If LCase(PrimaParola(Text)) = "alias" Then
        X.decAlias = PrimaParola(RestoStringa(Text))
    Else
        X.decAlias = ""
    End If
    
    'Aggiunta parametri
    
    p1 = InStr(1, Text, "(", vbTextCompare)
    If Mid(Text, Len(Text) - 2) = "()" Then
        p2 = InStrRev(Text, ")", Len(Text) - 2, vbTextCompare)
    Else
        p2 = InStrRev(Text, ")", -1, vbTextCompare)
    End If
    
    'Debug.Print Text
    f = Split(Mid(Text, p1 + 1, p2 - p1 - 1), ",")
    
    For i = LBound(f) To UBound(f)
        
        r = PrimaParola(f(i))
        Select Case LCase(r)
        Case "byval", "byref", "optional", "paramarray"
            Y.paramName = r & " " & PrimaParola(RestoStringa(f(i)))
            r = RestoStringa(RestoStringa(f(i)))
        Case Else
            Y.paramName = Trim(PrimaParola(f(i)))
            r = RestoStringa(f(i))
        End Select
                
        Y.paramType = RestoStringa(r)
        
        'Add a new key
        Y.idKey = "par" & X.decParamsID
        X.decParamsID = X.decParamsID + 1
        
        X.decParams.Add Y, Y.idKey
        Set Y = Nothing
    Next i
    
    If Not X.decSub Then _
        X.decReturnType = PrimaParola(RestoStringa(Mid(Text, p2 + 1)))
    
    'Add a new key
    X.idKey = "dec" & Api.cDeclaresID
    Api.cDeclaresID = Api.cDeclaresID + 1
    
    'Aggiunge all'insieme
    
    Api.cDeclares.Add X, X.idKey
End Sub


Sub AddDep1(d As apiDeclares, view As clsAPI, Api As clsAPI)
    
    Dim p As apiParams
    Dim t As apiType, t2 As apiType
    Dim ext As Boolean
    
    For Each p In d.decParams
    
        Select Case LCase(Trim(p.paramType))
        Case "boolean", "byte", "integer", "long", "single", "double", "string", "variant", "any", "currency", "decimal", "date", "object"
            DoEvents
            'Do nothing...
        Case Else
            'It's a defined type
            For Each t In Api.cTypes
                If LCase(Trim(t.decName)) = LCase(Trim(p.paramType)) Then
                    
                    ext = False
                    For Each t2 In view.cTypes
                        If LCase(Trim(t2.decName)) = LCase(Trim(t.decName)) Then ext = True
                    Next t2

                    If Not ext Then
                        view.cTypes.Add t, t.idKey
                        Call AddDep2(t, view, Api)
                    End If
                End If
            Next t
        End Select
    Next p
End Sub

Sub AddDep2(t As apiType, view As clsAPI, Api As clsAPI)

    Dim p As apiParams
    Dim t2 As apiType

    For Each p In t.decParams
    
        Select Case LCase(Trim(p.paramType))
        Case "boolean", "byte", "integer", "long", "single", "double", "string", "variant", "any", "currency", "decimal", "date", "object"
            DoEvents
            'Do nothing...
        Case Else
            'It's a defined type
            For Each t2 In Api.cTypes
                If LCase(Trim(t2.decName)) = LCase(Trim(p.paramName)) Then
                    view.cTypes.Add t2, t2.idKey
                    Call AddDep2(t2, view, Api)
                End If
            Next t2
        End Select
    Next p
End Sub

Sub RemDep(Key As String, view As clsAPI)

    Dim d As apiDeclares
    Dim t As apiType, t2 As apiType
    Dim p As apiParams

    If Left(Key, 1) = "d" Then
        
        Set d = view.cDeclares(Key)
        
        For Each p In d.decParams
            
            Select Case LCase(Trim(p.paramType))
            Case "boolean", "byte", "integer", "long", "single", "double", "string", "variant", "any", "currency", "decimal", "date", "object"
                DoEvents
                'Do nothing...
            Case Else
                'It's a defined type
                For Each t In view.cTypes
                    If LCase(Trim(t.decName)) = LCase(Trim(p.paramName)) Then
                        Call RemDep(t.idKey, view)
                        view.cTypes.Remove t2.idKey
                    End If
                Next t
            End Select

        Next p
        
        view.cDeclares.Remove Key
        
    Else
    
        Set t = view.cTypes(Key)
        
        For Each p In t.decParams
            
            Select Case LCase(Trim(p.paramType))
            Case "boolean", "byte", "integer", "long", "single", "double", "string", "variant", "any", "currency", "decimal", "date", "object"
                DoEvents
                'Do nothing...
            Case Else
                'It's a defined type
                For Each t2 In view.cTypes
                    If LCase(Trim(t2.decName)) = LCase(Trim(p.paramName)) Then
                        Call RemDep(t2.idKey, view)
                        view.cTypes.Remove t2.idKey
                    End If
                Next t2
            End Select

        Next p
        
        view.cTypes.Remove Key
    End If
End Sub
