Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim FileName As String
    FileName = App.Path & "\data files\logs\errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub


Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
On Error Resume Next
    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal index As Long, ByVal Text As String, ByVal FN As String)
    Dim FileName As String
    Dim F As Long

    If ServerLog Then
        If Not LPE(index) Then
            FileName = App.Path & "\data\logs\" & FN
    
            If Not FileExist(FileName, True) Then
                F = FreeFile
                Open FileName For Output As #F
                Close #F
            End If
    
            F = FreeFile
            Open FileName For Append As #F
            Print #F, Now & ": " & Text
            Close #F
        End If
    End If

End Sub

Sub AddLog2(ByVal Text As String, ByVal FN As String)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\data\logs\" & FN

    If Not FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    F = FreeFile
    Open FileName For Append As #F
    Print #F, Now & ": " & Text
    Close #F

End Sub

' gets a string from a text file
Public Function GetVar(File As String, header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        'FileExist = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
        If (dir(App.Path & "\" & FileName)) <> "" Then
            FileExist = True
        End If

    Else
        'FileExist = (GetAttr(FileName) And vbDirectory) = 0
        If (dir(FileName)) <> "" Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    PutVar App.Path & "\data\options.ini", "OPTIONS", "DisableAdmins", Options.DisableAdmins
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Update", Options.Update
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Instructions", Options.Instructions
    PutVar App.Path & "\data\options.ini", "OPTIONS", "ExpMultiplier", STR(Options.ExpMultiplier)
End Sub

Public Sub LoadOptions()
    
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Website")
    Options.DisableAdmins = GetVar(App.Path & "\data\options.ini", "OPTIONS", "DisableAdmins")
    Options.Update = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Update")
    Options.Instructions = GetVar(App.Path & "\data\options.ini", "OPTIONS", "instructions")
    Options.ExpMultiplier = GetVar(App.Path & "\data\options.ini", "OPTIONS", "ExpMultiplier")

End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim FileName As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    FileName = App.Path & "\data\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open FileName For Append As #F
    Print #F, IP & "," & GetPlayerName(BannedByIndex)
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned by " & GetPlayerName(BannedByIndex) & "!", White, False, True)
    Call AddLog(BannedByIndex, GetPlayerName(BannedByIndex) & " ha baneado a " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "Has sido baneado por " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
If frmServer.chkTroll.Value = vbChecked Then Exit Sub
    Dim FileName As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    FileName = App.Path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open FileName For Append As #F
    Print #F, IP & "," & "Server"
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned by the Server" & "!", White, False, True)
    Call AddLog(0, "The Server" & " ha baneado a " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "Has sido baneado por " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim FileName As String
    FileName = "data\accounts\" & Trim$(Name) & ".bin"

    If FileExist(FileName) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal password As String) As Boolean
    Dim FileName As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        FileName = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH + 1, RightPassword
        Close #nFileNum
        
        RightPassword = DesEncriptatePassword(RightPassword)

        If UCase$(Trim$(password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal password As String)
    Dim i As Long
    
    ClearPlayer index
    
    player(index).login = Name
    player(index).password = EncriptatePassword(password)

    Call SavePlayer(index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long) As Boolean

    If LenB(Trim$(player(index).Name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim F As Long
    Dim N As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(player(index).Name)) = 0 Then
        
        spritecheck = False
        
        player(index).Name = Name
        player(index).Sex = Sex
        player(index).Class = ClassNum
        
        If player(index).Sex = SEX_MALE Then
            player(index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            player(index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        player(index).level = 1

        For N = 1 To Stats.Stat_Count - 1
            player(index).stat(N) = Class(ClassNum).stat(N)
        Next N

        player(index).dir = DIR_DOWN
        player(index).map = Class(player(index).Class).StartMap
        player(index).X = Class(player(index).Class).StartMapX
        player(index).Y = Class(player(index).Class).StartMapY
        player(index).dir = DIR_DOWN
        player(index).vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        player(index).vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)
        
        'Kill Counter
        player(index).Kill = 0
        player(index).Dead = 0
        player(index).NpcKill = 0
        player(index).NpcDead = 0
        player(index).EnviroDead = 0
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For N = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(N) > 0 Then
                    ' item exist?
                    If Len(Trim$(item(Class(ClassNum).StartItem(N)).Name)) > 0 Then
                        player(index).Inv(N).Num = Class(ClassNum).StartItem(N)
                        player(index).Inv(N).Value = Class(ClassNum).StartValue(N)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For N = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(N) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartSpell(N)).Name)) > 0 Then
                        player(index).Spell(N) = Class(ClassNum).StartSpell(N)
                    End If
                End If
            Next
        End If
        
        ' set money bags
        player(index).RupeeBags = INITIAL_BAGS
        
        'set the inventory weight
        Call SetPlayerMaxWeight(index, INITIAL_MAX_WEIGHT)
        
        ' Append name to file
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function
' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal index As Long)
    Dim FileName As String
    Dim F As Long
    
    If IsServerBug Then Exit Sub

    FileName = App.Path & "\data\accounts\" & Trim$(player(index).login) & ".bin"
    
    F = FreeFile
    
    Open FileName For Binary As #F
    Put #F, , player(index)
    Close #F
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
    Dim FileName As String
    Dim F As Long
    Call ClearPlayer(index)
    FileName = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , player(index)
    Close #F
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index)), LenB(TempPlayer(index)))
    Set TempPlayer(index).buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(player(index)), LenB(player(index)))
    player(index).login = vbNullString
    player(index).password = vbNullString
    player(index).Name = vbNullString
    player(index).Class = 1

    frmServer.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(3) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim FileName As String
    Dim File As String
    FileName = App.Path & "\data\classes.ini"
    Max_Classes = 3

    If Not FileExist(FileName, True) Then
        File = FreeFile
        Open FileName For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim i As Long, N As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim X As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        FileName = App.Path & "\data\classes.ini"
        Max_Classes = val(GetVar(FileName, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).TranslatedName = GetVar(FileName, "CLASS" & i, "TranslatedName")
        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For N = 0 To UBound(tmpArray)
            Class(i).MaleSprite(N) = val(tmpArray(N))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For N = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(N) = val(tmpArray(N))
        Next
        
        ' continue
        Class(i).Face = val(GetVar(FileName, "CLASS" & i, "Face"))
        Class(i).stat(Stats.Strength) = val(GetVar(FileName, "CLASS" & i, "Strength"))
        Class(i).stat(Stats.Endurance) = val(GetVar(FileName, "CLASS" & i, "Endurance"))
        Class(i).stat(Stats.Intelligence) = val(GetVar(FileName, "CLASS" & i, "Intelligence"))
        Class(i).stat(Stats.Agility) = val(GetVar(FileName, "CLASS" & i, "Agility"))
        Class(i).stat(Stats.willpower) = val(GetVar(FileName, "CLASS" & i, "Willpower"))
        Class(i).StartMap = val(GetVar(FileName, "CLASS" & i, "StartMap"))
        Class(i).StartMapX = val(GetVar(FileName, "CLASS" & i, "StartMapX"))
        Class(i).StartMapY = val(GetVar(FileName, "CLASS" & i, "StartMapY"))
        
        
        ' how many starting items?
        startItemCount = val(GetVar(FileName, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For X = 1 To startItemCount
                Class(i).StartItem(X) = val(GetVar(FileName, "CLASS" & i, "StartItem" & X))
                Class(i).StartValue(X) = val(GetVar(FileName, "CLASS" & i, "StartValue" & X))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = val(GetVar(FileName, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For X = 1 To startSpellCount
                Class(i).StartSpell(X) = val(GetVar(FileName, "CLASS" & i, "StartSpell" & X))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim i As Long
    Dim X As Long
    
    FileName = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "TranslatedName", Trim$(Class(i).TranslatedName))
        Call PutVar(FileName, "CLASS" & i, "Maleprite", "1")
        Call PutVar(FileName, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(FileName, "CLASS" & i, "Face", "1")
        Call PutVar(FileName, "CLASS" & i, "Strength", STR(Class(i).stat(Stats.Strength)))
        Call PutVar(FileName, "CLASS" & i, "Endurance", STR(Class(i).stat(Stats.Endurance)))
        Call PutVar(FileName, "CLASS" & i, "Intelligence", STR(Class(i).stat(Stats.Intelligence)))
        Call PutVar(FileName, "CLASS" & i, "Agility", STR(Class(i).stat(Stats.Agility)))
        Call PutVar(FileName, "CLASS" & i, "Willpower", STR(Class(i).stat(Stats.willpower)))
        Call PutVar(FileName, "CLASS" & i, "StartMap", STR(Class(i).StartMap))
        Call PutVar(FileName, "CLASS" & i, "StartMapX", STR(Class(i).StartMapX))
        Call PutVar(FileName, "CLASS" & i, "StartMapY", STR(Class(i).StartMapY))
        
        ' loop for items & values
        For X = 1 To UBound(Class(i).StartItem)
            Call PutVar(FileName, "CLASS" & i, "StartItem" & X, STR(Class(i).StartItem(X)))
            Call PutVar(FileName, "CLASS" & i, "StartValue" & X, STR(Class(i).StartValue(X)))
        Next
        ' loop for spells
        For X = 1 To UBound(Class(i).StartSpell)
            Call PutVar(FileName, "CLASS" & i, "StartSpell" & X, STR(Class(i).StartSpell(X)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim FileName As String
    FileName = App.Path & "\data\classes.ini"

    If Not FileExist(FileName, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim FileName As String
    Dim F  As Long
    FileName = App.Path & "\data\items\item" & ItemNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        FileName = App.Path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , item(i)
        Close #F
        If Trim$(Replace(item(i).TranslatedName, vbNullChar, "")) = "" Then item(i).TranslatedName = GetTranslation(item(i).Name)
    Next

End Sub

Sub LoadItem(i As Long)
    Dim FileName As String
    Dim F As Long

        FileName = App.Path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , item(i)
        Close #F
        If Trim$(Replace(item(i).TranslatedName, vbNullChar, "")) = "" Then item(i).TranslatedName = GetTranslation(item(i).Name)

Call SendUpdateItemToAll(i)
End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(item(index)), LenB(item(index)))
    item(index).Name = vbNullString
    item(index).Desc = vbNullString
    item(index).Sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopnum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\shops\shop" & shopnum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Shop(shopnum)
    Close #F
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        FileName = App.Path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Shop(i)
        Close #F
        If Trim$(Replace(Shop(i).TranslatedName, vbNullChar, "")) = "" Then Shop(i).TranslatedName = GetTranslation(Shop(i).Name)
    Next

End Sub


Sub LoadShop(i As Long)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\data\shops\shop" & i & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Shop(i)
    Close #F
    If Trim$(Replace(Shop(i).TranslatedName, vbNullChar, "")) = "" Then Shop(i).TranslatedName = GetTranslation(Shop(i).Name)

    Call SendUpdateShopToAll(i)
    
End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellnum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\spells\spells" & spellnum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Spell(spellnum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Guardando hechizos... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        FileName = App.Path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Spell(i)
        Close #F
        If Trim$(Replace(Spell(i).TranslatedName, vbNullChar, "")) = "" Then Spell(i).TranslatedName = GetTranslation(Spell(i).Name)
    Next

End Sub

Sub LoadSpell(i As Long)
    Dim FileName As String
    Dim F As Long
    
    FileName = App.Path & "\data\spells\spells" & i & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Spell(i)
    Close #F
    If Trim$(Replace(Spell(i).TranslatedName, vbNullChar, "")) = "" Then Spell(i).TranslatedName = GetTranslation(Spell(i).Name)
    
    Call SendUpdateSpellToAll(i)
End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal npcnum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\npcs\npc" & npcnum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , NPC(npcnum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        FileName = App.Path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , NPC(i)
        Close #F
        If Trim$(Replace(NPC(i).TranslatedName, vbNullChar, "")) = "" Then NPC(i).TranslatedName = GetTranslation(NPC(i).Name)
    Next

End Sub

Sub LoadNpc(i As Long)
    Dim FileName As String
    Dim F As Long

        FileName = App.Path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , NPC(i)
        Close #F
        If Trim$(Replace(NPC(i).TranslatedName, vbNullChar, "")) = "" Then NPC(i).TranslatedName = GetTranslation(NPC(i).Name)

Call SendUpdateNpcToAll(i)
End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(index)), LenB(NPC(index)))
    NPC(index).Name = vbNullString
    NPC(index).TranslatedName = vbNullString
    NPC(index).AttackSay = vbNullString
    NPC(index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        FileName = App.Path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Resource(i)
        Close #F
        If Trim$(Replace(Resource(i).TranslatedName, vbNullChar, "")) = "" Then Resource(i).TranslatedName = GetTranslation(Resource(i).Name)
        
    Next

End Sub

Sub LoadResource(i As Long)
    Dim FileName As String
    Dim F As Long

        FileName = App.Path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Resource(i)
        Close #F
        If Trim$(Replace(Resource(i).TranslatedName, vbNullChar, "")) = "" Then Resource(i).TranslatedName = GetTranslation(Resource(i).Name)

Call SendUpdateResourceToAll(i)
End Sub


Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        FileName = App.Path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Animation(i)
        Close #F
        If Trim$(Replace(Animation(i).TranslatedName, vbNullChar, "")) = "" Then Animation(i).TranslatedName = GetTranslation(Animation(i).Name)
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).Name = vbNullString
    Animation(index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal mapnum As Long, ByRef map As MapRec)
    Dim FileName As String
    Dim F As Long
    Dim X As Long
    Dim Y As Long
    Dim Data() As Byte
    Data = GetMapData(map)
    FileName = App.Path & "\data\maps\map" & mapnum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Compress(Data)
    Close #F
End Sub

Sub LoadMaps()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Dim X As Long
    Dim Y As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        FileName = App.Path & "\data\maps\map" & i & ".dat"
    
        Dim CompressedData() As Byte
        CompressedData = ReadFile(FileName)
        MapCache(i).Data = CompressedData

        Call SetServerMapData(i, Decompress(CompressedData))
        ClearTempTile i
        CacheResources i
        DoEvents
    Next

End Sub

Sub LoadMap(i As Long)
    Dim FileName As String
    Dim F As Long
    Dim X As Long
    Dim Y As Long
    Call CheckMaps

        FileName = App.Path & "\data\maps\map" & i & ".dat"
    
        Dim CompressedData() As Byte
        CompressedData = ReadFile(FileName)
        MapCache(i).Data = CompressedData

        Call SetServerMapData(i, Decompress(CompressedData))
        ClearTempTile i
        CacheResources i

End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Dim TestMap As MapRec
            Call SaveMap(i, TestMap)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal index As Long, ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(mapnum, index)), LenB(MapItem(mapnum, index)))
    MapItem(mapnum, index).playerName = vbNullString
    
    CheckMapItemHighIndex mapnum, index, False
    
End Sub

Sub ClearMapItems()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal mapnum As Long)
    ReDim MapNpc(mapnum).NPC(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapnum).NPC(index)), LenB(MapNpc(mapnum).NPC(index)))
End Sub

Sub ClearMapNpcs()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next
    Next

End Sub

Sub BackupMap(ByVal mapnum As Long, ByVal Revision As Long)
On Error Resume Next
Dim FilePath As String
Dim BackupPath As String
Dim FileName As String
Dim NewFileName As String

FilePath = App.Path & "\data\maps\"
FileName = FilePath & "map" & mapnum & ".dat"
BackupPath = FilePath & "revisions\"
MkDir BackupPath
NewFileName = BackupPath & "map" & mapnum & "-" & Revision & ".dat"
FileCopy FileName, NewFileName
End Sub

Sub ClearMap(ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(map(mapnum)), LenB(map(mapnum)))
    
    map(mapnum).Name = vbNullString
    map(mapnum).MaxX = MAX_MAPX
    map(mapnum).MaxY = MAX_MAPY
    ReDim map(mapnum).Tile(0 To map(mapnum).MaxX, 0 To map(mapnum).MaxY)
    ' Reset the values for if a player is on the map or not
    ClearMapReference mapnum
    ' Reset the map cache array for this map.
    MapCache(mapnum).Data = vbNullString
    
    'reset tempmap data
    Call ZeroMemory(ByVal VarPtr(TempMap(mapnum)), LenB(TempMap(mapnum)))
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).TranslatedName)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal vital As Vitals) As Long
    Select Case vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal stat As Stats) As Long
    GetClassStat = Class(ClassNum).stat(stat)
End Function

Sub SaveBank(ByVal index As Long)
    Dim FileName As String
    Dim F As Long
    
    If IsServerBug Then Exit Sub
    
    FileName = App.Path & "\data\banks\" & Trim$(player(index).login) & ".bin"
    
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Bank(index)
    Close #F
End Sub

Public Sub LoadBank(ByVal index As Long, ByVal Name As String)
    Dim FileName As String
    Dim F As Long

    Call ClearBank(index)

    FileName = App.Path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(FileName, True) Then
        Call SaveBank(index)
        Exit Sub
    End If

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Bank(index)
    Close #F

End Sub

Sub ClearBank(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(index)), LenB(Bank(index)))
End Sub

Sub ClearParty(ByVal partynum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partynum)), LenB(Party(partynum)))
End Sub

Sub ClearProjectile(ByVal index As Long, ByVal PlayerProjectile As Long)
    ' clear the projectile
    With TempPlayer(index).ProjecTile(PlayerProjectile)
        .Direction = 0
        .Pic = 0
        .TravelTime = 0
        .X = 0
        .Y = 0
        .range = 0
        .Damage = 0
        .Speed = 0
    End With
End Sub

' ***********
' ** Doors **
' ***********

Sub SaveDoors()
    Dim i As Long

    For i = 1 To MAX_DOORS
        Call SaveDoor(i)
    Next

End Sub

Sub SaveDoor(ByVal DoorNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\doors\door" & DoorNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Doors(DoorNum)
    Close #F
End Sub

Sub LoadDoors()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckDoors

    For i = 1 To MAX_DOORS
        FileName = App.Path & "\data\doors\door" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Doors(i)
        Close #F
        If Trim$(Replace(Doors(i).TranslatedName, vbNullChar, "")) = "" Then Doors(i).TranslatedName = GetTranslation(Doors(i).Name)
    Next

End Sub

Sub CheckDoors()
    Dim i As Long

    For i = 1 To MAX_DOORS
        If Not FileExist("\Data\doors\door" & i & ".dat") Then
            Call SaveDoor(i)
        End If
    Next

End Sub

Sub ClearDoor(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Doors(index)), LenB(Doors(index)))
    Doors(index).Name = vbNullString
    ReDim Door(index)
End Sub

Sub ClearDoors()
    Dim i As Long

    For i = 1 To MAX_DOORS
        Call ClearDoor(i)
    Next

End Sub

Sub Savemovements()
    Dim i As Long

    For i = 1 To MAX_MOVEMENTS
        Call Savemovement(i)
    Next

End Sub

Sub Savemovement(ByVal MovementNum As Long)
    Dim FileName As String
    Dim F As Long, i As Byte
    FileName = App.Path & "\data\movements\movement" & MovementNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Movements(MovementNum).Name
        Put #F, , Movements(MovementNum).Type
        Put #F, , Movements(MovementNum).MovementsTable.Actual
        Put #F, , Movements(MovementNum).MovementsTable.nelem
        If Movements(MovementNum).MovementsTable.nelem > 0 Then
            ReDim Preserve Movements(MovementNum).MovementsTable.vect(1 To Movements(MovementNum).MovementsTable.nelem)
            For i = 1 To Movements(MovementNum).MovementsTable.nelem
                Put #F, , Movements(MovementNum).MovementsTable.vect(i).Data.Direction
                Put #F, , Movements(MovementNum).MovementsTable.vect(i).Data.NumberOfTiles
            Next
        End If
        Put #F, , Movements(MovementNum).Repeat
    Close #F
End Sub

Sub Loadmovements()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Dim j As Byte
    
    Call Checkmovements

    For i = 1 To MAX_MOVEMENTS
        FileName = App.Path & "\data\movements\movement" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Movements(i).Name
            Get #F, , Movements(i).Type
            Get #F, , Movements(i).MovementsTable.Actual
            Get #F, , Movements(i).MovementsTable.nelem
            If Movements(i).MovementsTable.nelem > 0 Then
                ReDim Movements(i).MovementsTable.vect(1 To Movements(i).MovementsTable.nelem)
                For j = 1 To Movements(i).MovementsTable.nelem
                    Get #F, , Movements(i).MovementsTable.vect(j).Data.Direction
                    Get #F, , Movements(i).MovementsTable.vect(j).Data.NumberOfTiles
                Next
            End If
            Get #F, , Movements(i).Repeat
        Close #F
    Next

End Sub

Sub Checkmovements()
    Dim i As Long

    For i = 1 To MAX_MOVEMENTS
        If Not FileExist("\Data\movements\movement" & i & ".dat") Then
            Call Savemovement(i)
        End If
    Next

End Sub

Sub Clearmovement(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Movements(index)), LenB(Movements(index)))
    Movements(index).Name = vbNullString
    ReDim Movement(index)
End Sub

Sub Clearmovements()
    Dim i As Long

    For i = 1 To MAX_MOVEMENTS
        Call Clearmovement(i)
    Next

End Sub

Sub SaveActions()
    Dim i As Long

    For i = 1 To MAX_ACTIONS
        Call SaveAction(i)
    Next

End Sub

Sub SaveAction(ByVal ActionNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\Actions\Action" & ActionNum & ".dat"
    F = FreeFile
    'Actions(ActionNum).TranslatedName = GetTranslation(Actions(ActionNum).Name)
    Open FileName For Binary As #F
        Put #F, , Actions(ActionNum).Name
        Put #F, , Actions(ActionNum).Type
        Put #F, , Actions(ActionNum).Moment
        Put #F, , Actions(ActionNum).Data1
        Put #F, , Actions(ActionNum).Data2
        Put #F, , Actions(ActionNum).Data3
        Put #F, , Actions(ActionNum).Data4
    Close #F
    
End Sub

Sub LoadActions()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckActions

    For i = 1 To MAX_ACTIONS
        FileName = App.Path & "\data\Actions\Action" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Actions(i).Name
            Get #F, , Actions(i).Type
            Get #F, , Actions(i).Moment
            Get #F, , Actions(i).Data1
            Get #F, , Actions(i).Data2
            Get #F, , Actions(i).Data3
            Get #F, , Actions(i).Data4

        Close #F
        If Trim$(Replace(Actions(i).TranslatedName, vbNullChar, "")) = "" Then Actions(i).TranslatedName = GetTranslation(Actions(i).Name)
        
    Next

End Sub

Sub CheckActions()
    Dim i As Long

    For i = 1 To MAX_ACTIONS
        If Not FileExist("\Data\Actions\Action" & i & ".dat") Then
            Call SaveAction(i)
        End If
    Next

End Sub

Sub ClearAction(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Actions(index)), LenB(Actions(index)))
    Actions(index).Name = vbNullString
End Sub

Sub ClearActions()
    Dim i As Long

    For i = 1 To MAX_ACTIONS
        Call ClearAction(i)
    Next

End Sub

Sub SaveCustomSprites()
    Dim i As Long

    For i = 1 To MAX_CUSTOM_SPRITES
        Call SaveCustomSprite(i)
    Next

End Sub

Sub SaveCustomSprite(ByVal CustomSpriteNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\CustomSprites\CustomSprite" & CustomSpriteNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
            'Put #F, , CustomSprites(CustomSpriteNum)
            'Open FileName For Binary As #F
    Dim Data() As Byte
    Put #F, , GetCustomSpriteData(CustomSpriteNum)
            'With CustomSprites(CustomSpriteNum)
                'Put #F, , .Name
                'Put #F, , .NLayers
                'Dim i As Byte
                'For i = 1 To .NLayers
                    'Put #F, , .Layers(i).Sprite
                    'Put #F, , .Layers(i).UseCenterPosition
                    'Put #F, , .Layers(i).UsePlayerSprite
                    'Dim j As Byte, k As Byte
                    'For j = 0 To MAX_DIRECTIONS - 1
                        'For k = 0 To max_anims - 1
                            'Put #F, , .Layers(i).fixed.EnabledAnims(j, k)
                        'Next
                    'Next
                    'For j = 0 To MAX_DIRECTIONS - 1
                        'Put #F, , .Layers(i).CentersPositions(j).X
                        'Put #F, , .Layers(i).CentersPositions(j).Y
                    'Next
                'Next
                
                   
        'End With
    Close #F
End Sub




Sub LoadCustomSprites()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckCustomSprites

    For i = 1 To MAX_CUSTOM_SPRITES
        FileName = App.Path & "\data\CustomSprites\CustomSprite" & i & ".dat"
        'F = FreeFile
        'Open FileName For Binary As #F
        
        Dim Data() As Byte
        Data = ReadFile(FileName)
        'Get #F, , Data
        
        Call SetCustomSpriteData(i, Data)
        
            'Get #F, , CustomSprites(i)
            'With CustomSprites(i)
            'Get #F, , .Name
            'Get #F, , .NLayers
            'If .NLayers <> 0 Then
                'ReDim .Layers(1 To .NLayers)
            'End If
            'Dim j As Byte
            'For j = 1 To .NLayers
                'Get #F, , .Layers(j).Sprite
                'Get #F, , .Layers(j).UseCenterPosition
                'Get #F, , .Layers(j).UsePlayerSprite
                'Dim k As Byte, l As Byte
                'For j = 0 To MAX_DIRECTIONS - 1
                        'For k = 0 To max_anims - 1
                            'Put #F, , .Layers(i).fixed.EnabledAnims(j, k)
                        'Next
                    'Next
                'For k = 0 To MAX_DIRECTIONS - 1
                    'Get #F, , .Layers(j).CentersPositions(k).X
                    'Get #F, , .Layers(j).CentersPositions(k).Y
                'Next
            'Next
                
                   
        'End With
        'Close #F
    Next

End Sub

Sub CheckCustomSprites()
    Dim i As Long

    For i = 1 To MAX_CUSTOM_SPRITES
        If Not FileExist("\Data\CustomSprites\CustomSprite" & i & ".dat") Then
            Call SaveCustomSprite(i)
        End If
    Next

End Sub

Sub ClearCustomSprite(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(CustomSprites(index)), LenB(CustomSprites(index)))
    CustomSprites(index).Name = vbNullString
End Sub

Sub ClearCustomSprites()
    Dim i As Long

    For i = 1 To MAX_CUSTOM_SPRITES
        Call ClearCustomSprite(i)
    Next

End Sub

Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
ErrorHandler:
' if an error occurs, this function returns False
End Function

Sub ClearSingleMapNpc(ByVal index As Long, ByVal mapnum As Long)

    DeleteNPCFromMapRef mapnum, index
    
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapnum).NPC(index)), LenB(MapNpc(mapnum).NPC(index)))

    If index >= TempMap(mapnum).npc_highindex Then
        Call SetMapNPCHighIndex(mapnum, index)
    End If
    

End Sub


Sub SavePets()
    Dim i As Long

    For i = 1 To MAX_PETS
        Call SavePet(i)
    Next

End Sub

Sub SavePet(ByVal PetNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\Pets\Pet" & PetNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Pet(PetNum).Name
        Put #F, , Pet(PetNum).npcnum
        Put #F, , Pet(PetNum).TamePoints
        Put #F, , Pet(PetNum).ExpProgression
        Put #F, , Pet(PetNum).pointsprogression
        Put #F, , Pet(PetNum).MaxLevel
    Close #F
End Sub

Sub LoadPets()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    
    Call CheckPets

    For i = 1 To MAX_PETS
        FileName = App.Path & "\data\Pets\Pet" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Pet(i).Name
            Get #F, , Pet(i).npcnum
            Get #F, , Pet(i).TamePoints
            Get #F, , Pet(i).ExpProgression
            Get #F, , Pet(i).pointsprogression
            Get #F, , Pet(i).MaxLevel

        Close #F
    Next

End Sub

Sub LoadPet(i As Long)
    Dim FileName As String
    Dim F As Long

        FileName = App.Path & "\data\Pets\Pet" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Pet(i).Name
            Get #F, , Pet(i).npcnum
            Get #F, , Pet(i).TamePoints
            Get #F, , Pet(i).ExpProgression
            Get #F, , Pet(i).pointsprogression
            Get #F, , Pet(i).MaxLevel
        Close #F

Call SendUpdatePetToAll(i)

End Sub

Sub CheckPets()
    Dim i As Long

    For i = 1 To MAX_PETS
        If Not FileExist("\Data\Pets\Pet" & i & ".dat") Then
            Call SavePet(i)
        End If
    Next

End Sub

Sub ClearPet(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Pet(index)), LenB(Pet(index)))
    Pet(index).Name = vbNullString
End Sub

Sub ClearPets()
    Dim i As Long

    For i = 1 To MAX_PETS
        Call ClearPet(i)
    Next

End Sub

Function GetAccountsPassword() As String
    Dim FileName As String
    Dim F As Long
    Dim s As String
    
    FileName = App.Path & "\data\passwords\accounts.txt"
    GetAccountsPassword = vbNullString
    
    ' Make sure the file exists
    If Not FileExist("data\passwords\accounts.txt") Then
        Exit Function
    End If

    F = FreeFile
    Open FileName For Input As #F

    Input #F, s
    GetAccountsPassword = DesEncriptatePassword(Trim$(s))

    Close #F
End Function

Function GetMakeAdminPassword() As String
    Dim FileName As String
    Dim F As Long
    Dim s As String
    
    FileName = App.Path & "\data\passwords\items.txt"
    GetMakeAdminPassword = vbNullString
    
    ' Make sure the file exists
    If Not FileExist("data\passwords\items.txt") Then
        Exit Function
    End If

    F = FreeFile
    Open FileName For Input As #F

    Input #F, s
    GetMakeAdminPassword = DesEncriptatePassword(Trim$(s))

    Close #F
End Function

Function GetBanPassword() As String
If frmServer.chkTroll.Value = vbChecked Then Exit Function
    Dim FileName As String
    Dim F As Long
    Dim s As String
    
    FileName = App.Path & "\data\passwords\ban.txt"
    GetBanPassword = vbNullString
    
    ' Make sure the file exists
    If Not FileExist("data\passwords\ban.txt") Then
        Exit Function
    End If

    F = FreeFile
    Open FileName For Input As #F

    Input #F, s
    GetBanPassword = DesEncriptatePassword(Trim$(s))

    Close #F
End Function

Sub UnlockAccount(ByVal login As String)
    
    Dim FileName As String
    Dim F As Long

    
    If Not AccountExist(login) Then Exit Sub
    
    ZeroMemory AuxPlayer, Len(AuxPlayer)
    
    FileName = App.Path & "\data\accounts\" & Trim$(login) & ".bin"
    
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , AuxPlayer
    Close #F
    
    'If AuxPlayer.AccountBlocked Then
        'AuxPlayer.AccountBlocked = False
        'F = FreeFile
        'Open FileName For Binary As #F
        'Put #F, , AuxPlayer
        'Close #F
    'End If
    
    
End Sub

Sub CreateFile(ByRef File As String)
    Dim F As Long
    F = FreeFile
    Open File For Output As #F
    Close #F
End Sub


Sub DeleteCharName(ByVal Name As String)
    Dim FileName As String
    FileName = "data\accounts\charlist.txt"
    If FileExist(FileName) Then
        Call DeleteName(Name)
    End If
End Sub

Sub AddLine(ByRef File As String, ByRef Line As String)
    Dim F As Long
    F = FreeFile
    Open File For Append As #F
    Print #F, Line
    Close #F
End Sub

Function LineExists(ByRef File As String, ByRef Line As String) As Boolean
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open File For Input As #F

    Do While Not EOF(F) And Not LineExists
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Line)) Then
            LineExists = True
        End If

    Loop

    Close #F
End Function


Sub DeleteLine(ByRef File As String, ByRef extension As String, ByRef Line As String)
    
    Dim f1 As Long, f2 As Long
    Dim TempFile As String
    Dim GetLine As String
    TempFile = Replace(File, extension, ".tmp")
    f1 = FreeFile
    
    Open File For Input As #f1
    f2 = FreeFile
    Open TempFile For Output As #f2

    While Not EOF(f1)
       Input #f1, GetLine
        If GetLine <> Line Then
            Print #f2, GetLine
        End If
    Wend
    
    Close #f1, #f2
    
    Kill File       ' Deletes the Original File
    Name TempFile As File ' Renames the New File
End Sub


   
