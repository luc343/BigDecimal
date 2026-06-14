VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BigOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'
'                  Class BigOption
'
' Copyright (c) Lucien Cinc 2025-26
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

#Const Stats = 1     '1 = true, 0 = false

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Const StepSize As Long = 16        'step size
Private Const InitialSize As Long = 31     '0 to 15 = 16 initial list size

Private List() As LongPtr                   'list of bigdecimals
Private Size As Long                        'current ubound of the array
Private Num As Long                         'current number of active regs

Private Big_TarScale As Long                'max precision for a BigDecimal
Private Big_Cache As Boolean                'cache switch

#If Stats Then
    Private MaxReg As Long                  'maximum register count
    Private RegCount As Long                'total register count
    Private UnregCount As Long              'total unregister count
#End If

Private Ln2_Cache_Buf As String             'cache for ln(2)
Private Ln2_Cache_Scale As Long             'scale for cache of ln(2)
Private Ln10_Cache_Buf As String            'cache for ln(10)
Private Ln10_Cache_Scale As Long            'scale for cache of ln(10)
Private Pi_Cache_Buf As String              'cache for Pi
Private Pi_Cache_Scale As Long              'scale for cache of Pi

Private OptionOutputDebug As Boolean        'switch debug on/off
Private OptionAutoRound As Boolean          'switch rounding on/off when return a string
Private OptionAddTrailingZeroes As Boolean  'switch trailing zeroes on/off when using round()

'######################### BigOption Methods #########################

Private Sub Class_Initialize()
    Size = InitialSize
    ReDim List(0 To Size)
    Num = 0

    Big_TarScale = 100
    Big_Cache = True

    BigOption.Ln2_Buf = "0.69314718055994530941723212145817656807550013436025525412068000949339362" & _
			 "1969694715605863326996418687542001481020570685733685520235758130557032670" & _
			 "7516350759619307275708283714351903070386238916734711233501153644979552391" & _
			 "2047517268157493206515552473413952588295045300709532636664265410423915781" & _
			 "4952043740430385500801944170641671518644712839968171784546957026271631064" & _
			 "5461502572074024816377733896385506952606683411372738737229289564935470257" & _
			 "626520988596932019650585547647033067936544325476327449512504060694381"

    BigOption.Ln10_Buf = "2.30258509299404568401799145468436420760110148862877297603332790096757260" & _
			 "9677352480235997205089598298341967784042286248633409525465082806756666287" & _
			 "3690987816894829072083255546808437998948262331985283935053089653777326288" & _
			 "4616336622228769821988674654366747440424327436515504893431493939147961940" & _
			 "4400222105101714174800368808401264708068556774321622835522011480466371565" & _
			 "9121373450747856947683463616792101806445070648000277502684916746550586856" & _
			 "935673420670581136429224554405758925724208241314695689016758940256776"

    BigOption.Pi_Buf = "3.14159265358979323846264338327950288419716939937510582097494459230781640" & _
			 "6286208998628034825342117067982148086513282306647093844609550582231725359" & _
			 "4081284811174502841027019385211055596446229489549303819644288109756659334" & _
			 "4612847564823378678316527120190914564856692346034861045432664821339360726" & _
			 "0249141273724587006606315588174881520920962829254091715364367892590360011" & _
			 "3305305488204665213841469519415116094330572703657595919530921861173819326" & _
			 "117931051185480744623799627495673518857527248912279381830119491298337"

    BigOption.Ln2_Scale = Len(BigOption.Ln2_Buf)
    BigOption.Ln10_Scale = Len(BigOption.Ln10_Buf)
    BigOption.Pi_Scale = Len(BigOption.Pi_Buf)

    OptionOutputDebug = False
    OptionAutoRound = True
    OptionAddTrailingZeroes = False
End Sub

Private Sub Class_Terminate()
    Erase List
End Sub

'***************************************
'
'         Register a BigDecimal
'
'***************************************

Friend Sub Register(ByRef Val As BigDecimal)
    Dim i As Long
    Dim Ptr As LongPtr

    Ptr = ObjPtr(Val)
    For i = 0 To Size
	If List(i) = 0 Then
	    List(i) = Ptr
	    Num = Num + 1
	    GoTo Done
	End If
    Next i

    Size = Size + StepSize
    ReDim Preserve List(0 To Size)

    List(Num) = Ptr
    Num = Num + 1
Done:
#If Stats Then
    RegCount = RegCount + 1
    If RegCount - UnregCount > MaxReg Then
	MaxReg = RegCount - UnregCount
    End If
#End If
End Sub

'***************************************
'
'        Unregister a BigDecimal
'
'***************************************

Friend Sub Unregister(ByRef Val As BigDecimal)
    Dim i As Long
    Dim Ptr As LongPtr

    Ptr = ObjPtr(Val)
    For i = 0 To Size
	If List(i) = Ptr Then
	    List(i) = 0
	    Num = Num - 1
#If Stats Then
	    UnregCount = UnregCount + 1
#End If
	    Exit Sub
	End If
    Next i

    Err.Raise 5, "BigDecimalRegistry", "Object pointer not found in registry."
End Sub

Private Function CastPtr(ByVal Ptr As LongPtr) As BigDecimal
    Dim Temp As BigDecimal
    Dim Zero As LongPtr

    CopyMemory Temp, Ptr, LenB(Ptr)     'copy the ptr to Temp

    Set CastPtr = Temp

    CopyMemory Temp, Zero, LenB(Zero)    'un-point Temp so VBA doesn't try to free the memory
End Function

'Public Function GetRefCount(ByVal Obj As IUnknown) As Long
'    Dim pUnk As LongPtr
'    Dim RefCount As Long
'
'    If Obj Is Nothing Then
'        GetRefCount = 0
'        Exit Function
'    End If
'
'    pUnk = ObjPtr(Obj)
'
'    #If Win64 Then
'        CopyMemory RefCount, ByVal (pUnk + 8), 4
'    #Else
'        CopyMemory RefCount, ByVal (pUnk + 4), 4
'    #End If
'
'    GetRefCount = RefCount - 1
'End Function

Public Sub Stats()
#If Stats Then
    Debug.Print
    Debug.Print
    Debug.Print "BigOption Stats"
    Debug.Print "-----------------"
    Debug.Print "Max Regs: " & MaxReg
    Debug.Print "    Regs: " & RegCount
    Debug.Print "  UnRegs: " & UnregCount
    Debug.Print "   Count: " & Num

    MaxReg = 0
    RegCount = 0
    UnregCount = 0
#End If
End Sub

'####################### BigOption Properties ########################

'Count
Public Property Get Count() As Long
    Count = Num
End Property

'Cache
Public Property Get Cache() As Boolean
    Cache = Big_Cache
End Property

Public Property Let Cache(Flag As Boolean)
    Big_Cache = Flag
End Property

'TarScale
Public Property Get TarScale() As Long
    TarScale = Big_TarScale
End Property

Public Property Let TarScale(Val As Long)
Dim Temp As BigDecimal
    Dim i As Long

    Big_TarScale = Val

    If Num = 0 Then Exit Property

    For i = 0 To Size
	If List(i) <> 0 Then
	    Set Temp = CastPtr(List(i))
	    Temp.TarScale = Big_TarScale
	    Set Temp = Nothing

	    Yield.Events
	End If
    Next i
End Property

'AddTrailingZeroes
Public Property Get AddTrailingZeroes() As Boolean
    AddTrailingZeroes = OptionAddTrailingZeroes
End Property

Public Property Let AddTrailingZeroes(Flag As Boolean)
    Dim Temp As BigDecimal
    Dim i As Long

    OptionAddTrailingZeroes = Flag

    If Num = 0 Then Exit Property

    For i = 0 To Size
	If List(i) <> 0 Then
	    Set Temp = CastPtr(List(i))
	    Temp.AddTrailingZeroes = OptionAddTrailingZeroes
	    Set Temp = Nothing

	    Yield.Events
	End If
    Next i
End Property

'AutoRound
Public Property Get AutoRound() As Boolean
    AutoRound = OptionAutoRound
End Property

Public Property Let AutoRound(Flag As Boolean)
    Dim Temp As BigDecimal
    Dim i As Long

    OptionAutoRound = Flag

    If Num = 0 Then Exit Property

    For i = 0 To Size
	If List(i) <> 0 Then
	    Set Temp = CastPtr(List(i))
	    Temp.AutoRound = OptionAutoRound
	    Set Temp = Nothing

	    Yield.Events
	End If
    Next i
End Property

'OutputDebug
Public Property Get OutputDebug() As Boolean
    OutputDebug = OptionOutputDebug
End Property

Public Property Let OutputDebug(Flag As Boolean)
    Dim Temp As BigDecimal
    Dim i As Long

    OptionOutputDebug = Flag

    If Num = 0 Then Exit Property

    For i = 0 To Size
	If List(i) <> 0 Then
	    Set Temp = CastPtr(List(i))
	    Temp.OutputDebug = OptionOutputDebug
	    Set Temp = Nothing

	    Yield.Events
	End If
    Next i
End Property

'Ln2_Buf
Friend Property Get Ln2_Buf() As String
    If Big_Cache Then
	Ln2_Buf = Ln2_Cache_Buf
    Else
	Ln2_Buf = "0.693147180559945"
    End If
End Property

Friend Property Let Ln2_Buf(Buf As String)
    If Big_Cache Then
	Ln2_Cache_Buf = Buf
    End If
End Property

'Ln2_Scale
Friend Property Get Ln2_Scale() As Long
    If Big_Cache Then
	Ln2_Scale = Ln2_Cache_Scale
    Else
	Ln2_Scale = 15
    End If
End Property

Friend Property Let Ln2_Scale(Val As Long)
    If Big_Cache Then
	Ln2_Cache_Scale = Val
    End If
End Property

'Ln10_Buf
Friend Property Get Ln10_Buf() As String
    If Big_Cache Then
	Ln10_Buf = Ln10_Cache_Buf
    Else
	Ln10_Buf = "2.302585092994046"
    End If
End Property

Friend Property Let Ln10_Buf(Buf As String)
    If Big_Cache Then
	Ln10_Cache_Buf = Buf
    End If
End Property

'Ln10_Scale
Friend Property Get Ln10_Scale() As Long
    If Big_Cache Then
	Ln10_Scale = Ln10_Cache_Scale
    Else
	Ln10_Scale = 15
    End If
End Property

Friend Property Let Ln10_Scale(Val As Long)
    If Big_Cache Then
	Ln10_Cache_Scale = Val
    End If
End Property

'Pi_Buf
Friend Property Get Pi_Buf() As String
    If Big_Cache Then
	Pi_Buf = Pi_Cache_Buf
    End If
End Property

Friend Property Let Pi_Buf(Buf As String)
    If Big_Cache Then
	Pi_Cache_Buf = Buf
    End If
End Property

'Pi_Scale
Friend Property Get Pi_Scale() As Long
    If Big_Cache Then
	Pi_Scale = Pi_Cache_Scale
    End If
End Property

Friend Property Let Pi_Scale(Val As Long)
    If Big_Cache Then
	Pi_Cache_Scale = Val
    End If
End Property
