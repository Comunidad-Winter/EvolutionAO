Attribute VB_Name = "m_Canjes"
Option Explicit

Public Const MAXCANJES As Integer = 10000
Private NumCanjes As Integer

Private Type Canjeo
    Cantidad As Integer
    ObjIndex As Integer
    Puntos As Integer
End Type

Public Canjes() As Canjeo

Sub LoadCanjesData()

    'Canjes
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Canjes."

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize("ItemsCanje.dat")

    NumCanjes = val(Leer.GetValue("INIT", "Cantidad"))

    If NumCanjes <> 0 Then
        ReDim Canjes(1 To NumCanjes) As Canjeo

        Dim i As Long

        For i = 1 To NumCanjes
            With Canjes(i)
                .Cantidad = val(Leer.GetValue("SHOP" & i, "Cantidad"))
                .ObjIndex = val(Leer.GetValue("SHOP" & i, "ObjIndex"))
                .Puntos = val(Leer.GetValue("SHOP" & i, "Puntos"))
            End With
        Next i
    End If

    Set Leer = Nothing

End Sub

Public Function GetCanje(ByVal UserIndex As Integer, _
                         ByVal Canjea As Byte, _
                         Optional ByRef StrError As String = vbNullString) As Boolean

    If Canjea = 0 Then Exit Function
    If Canjea > NumCanjes Then Exit Function

    Dim MiObj As Obj

    With UserList(UserIndex)

        MiObj.Amount = Canjes(Canjea).Cantidad
        MiObj.ObjIndex = Canjes(Canjea).ObjIndex

        If .Stats.PuntosCanje < Canjes(Canjea).Puntos Then
            StrError = "No tienes los puntos suficientes."
            Exit Function
        ElseIf Not MeterItemEnInventario(UserIndex, MiObj) Then
            StrError = "No puedes cargar más objetos."
            Exit Function
        End If

        .Stats.PuntosCanje = .Stats.PuntosCanje - Canjes(Canjea).Puntos
        'Call WritePuntos(uIndex)

    End With

    GetCanje = True

End Function
