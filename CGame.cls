VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbytPlayers As Byte
Private mbytX As Byte
Private mbytY As Byte
Private mbytTurn As Byte
Private mbytValue(1 To 16, 1 To 16) As Byte
Private mblnIsHuman(1 To 4) As Boolean

Public Property Get Players() As Byte
    Players = mbytPlayers
End Property

Public Property Let Players(ByVal bytNumber As Byte)
    mbytPlayers = bytNumber
End Property

Public Property Get MaxX() As Byte
    MaxX = mbytX
End Property

Public Property Let MaxX(ByVal bytX As Byte)
    mbytX = bytX
End Property

Public Property Get MaxY() As Byte
    MaxY = mbytY
End Property

Public Property Let MaxY(ByVal bytY As Byte)
    mbytY = bytY
End Property

Private Sub Class_Initialize()
    Dim bytCounterX As Byte, bytCounterY As Byte
    
    'initialize all grid points (16x16 max) to 0
    For bytCounterX = 1 To 16
        For bytCounterY = 1 To 16
            Value bytCounterX, bytCounterY, 0
        Next bytCounterY
    Next bytCounterX
    
    'initialize object properties
    mbytPlayers = 0
    mbytX = 0
    mbytY = 0
    mbytTurn = 1
End Sub

Public Property Get PlayerTurn() As Byte
    PlayerTurn = mbytTurn
End Property

Public Sub NextPlayer()
    mbytTurn = mbytTurn + 1
    If mbytTurn > mbytPlayers Then mbytTurn = 1
    If Score(mbytTurn) = 0 Then NextPlayer
End Sub

Public Function Value(ByVal X As Byte, ByVal Y As Byte, _
        Optional ByVal NewValue As Variant)
    If Not IsMissing(NewValue) Then mbytValue(X, Y) = CByte(NewValue)
    Value = mbytValue(X, Y)
End Function

Public Function Score(ByVal bytPlayer As Byte) As Byte
    Dim bytCounterX As Byte, bytCounterY As Byte
    Dim bytScore As Byte
    bytScore = 0
    
    For bytCounterX = 1 To mbytX
        For bytCounterY = 1 To mbytY
            If Value(bytCounterX, bytCounterY) = bytPlayer Then _
                    bytScore = bytScore + 1
        Next bytCounterY
    Next bytCounterX
    
    Score = bytScore
End Function

Public Property Get IsHuman(ByVal bytPlayer) As Boolean
    IsHuman = mblnIsHuman(bytPlayer)
End Property

Public Property Let IsHuman(ByVal bytPlayer, ByVal blnNewVal As Boolean)
    mblnIsHuman(bytPlayer) = blnNewVal
End Property
