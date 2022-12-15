Sub Proyecto()
Dim CantidadPasajeros As Integer
Dim FilaPasajero As Integer
Dim ColumnaPasajero As Integer
CantidadPasajeros = Range(Hoja2.Cells(1, 1), Hoja2.Cells(1, 1).End(xlDown)).Count - 1
Dim Largo As Integer
Dim LargoComproSilla As Integer
LargoComproSilla = Range(Hoja3.Cells(1, 25), Hoja3.Cells(1, 25).End(xlDown)).Count
Dim AleatorioFila As Integer
Dim AleatorioCol As Integer
Dim parar As Boolean
Dim Ganancias As Double
Dim PosiblesGanancias As Double
Dim Fila As Integer
Dim Columna As Integer
Dim SillaVacia As Integer
Dim Parar2 As Boolean
Dim RndCambioCol As Integer
Dim RndCambioFil As Integer
Dim ColCambio As Integer
Dim FilCambio As Integer
Dim ContadorPasajeros As Integer
ContadorPasajeros = 1
Range("y3:y200").ClearContents
Hoja3.Range("a3:h33").Interior.ColorIndex = 0
Hoja3.Range("a3:h33").Font.ColorIndex = 0
Hoja3.Range("a3:h33").Font.Bold = False

Dim HacerCambio As Boolean
HacerCambio = False
Parar2 = False
parar = False
'Borrar el avion para iniciar
Hoja3.Range("C2:H32").ClearContents
Hoja2.Range("b2:b200").Interior.ColorIndex = 0
Hoja3.Range("S6:T100").ClearContents
Hoja3.Range("P7").ClearContents
Hoja3.Range("c2:h33").Interior.ColorIndex = 0
               
'ubicar a la gente que compró silla antes de fecha limite

For i = 1 To CantidadPasajeros
LargoComproSilla = Range(Hoja3.Cells(1, 25), Hoja3.Cells(1, 25).End(xlDown)).Count

    If Hoja2.Cells(i + 1, 10) <> "NULL" Then
        If Hoja2.Cells(i + 1, 14) >= 2 Then
            FilaPasajero = Hoja2.Cells(i + 1, 11)
            ColumnaPasajero = Hoja2.Cells(i + 1, 13)
            Hoja3.Cells(FilaPasajero + 1, ColumnaPasajero + 2) = ContadorPasajeros
            ContadorPasajeros = ContadorPasajeros + 1
            Hoja3.Cells(LargoComproSilla + 1, 25) = Hoja2.Cells(i + 1, 2)
            Hoja3.Cells(FilaPasajero + 1, ColumnaPasajero + 2).Interior.ColorIndex = 7
            Hoja3.Cells(FilaPasajero + 1, ColumnaPasajero + 2).Font.ColorIndex = 19
            Hoja3.Cells(FilaPasajero + 1, ColumnaPasajero + 2).Font.Bold = True
            
                If ContadorPasajeros = 100 Then
                    ContadorPasajeros = 1
                End If
            Hoja2.Cells(i + 1, 2).Interior.ColorIndex = 5
            


'calcular ganancias
            If Hoja2.Cells(i + 1, 12) = "A" Then
               Ganancias = Ganancias + Hoja3.Cells(i + 1, 11)
            End If
            
            If Hoja2.Cells(i + 1, 12) = "B" Then
                Ganancias = Ganancias + Hoja3.Cells(i + 1, 12)
            End If
            
            If Hoja2.Cells(i + 1, 12) = "C" Then
               Ganancias = Ganancias + Hoja3.Cells(i + 1, 13)
            End If
            
            If Hoja2.Cells(i + 1, 12) = "D" Then
                Ganancias = Ganancias + Hoja3.Cells(i + 1, 14)
            End If
        
        
            If Hoja2.Cells(i + 1, 12) = "E" Then
               Ganancias = Ganancias + Hoja3.Cells(i + 1, 15)
            End If
            
            If Hoja2.Cells(i + 1, 12) = "F" Then
                Ganancias = Ganancias + Hoja3.Cells(i + 1, 16)
            End If
            
        End If
    End If
    
    If Hoja2.Cells(i + 1, 10) <> "NULL" Then
        If Hoja2.Cells(i + 1, 14) = 1 Or Hoja2.Cells(i + 1, 14) = 0 Then
            FilaPasajero = Hoja2.Cells(i + 1, 11)
            ColumnaPasajero = Hoja2.Cells(i + 1, 13)
            Hoja3.Cells(FilaPasajero + 1, ColumnaPasajero + 2) = ContadorPasajeros
                ContadorPasajeros = ContadorPasajeros + 1
                If ContadorPasajeros = 100 Then
                    ContadorPasajeros = 1
                End If
            Hoja2.Cells(i + 1, 2).Interior.ColorIndex = 5
            
        End If
    End If
Next i

' Dejar libres las sillas que toca dejar libres
For i = 1 To 3
    Cells(33, 5 + i) = "X"
    
Next i

'Hacer el constructivo. Asignar aleatoreamente a todas las personas en una silla

For j = 1 To CantidadPasajeros
    If Hoja2.Cells(j + 1, 10) = "NULL" Then
        parar = False
        While parar = falso
        AleatorioCol = WorksheetFunction.RandBetween(1, 6)
        AleatorioFila = WorksheetFunction.RandBetween(1, 32)
            
            If Hoja3.Cells(AleatorioFila + 1, AleatorioCol + 2) = "" Then
                Hoja3.Cells(AleatorioFila + 1, AleatorioCol + 2) = ContadorPasajeros
                ContadorPasajeros = ContadorPasajeros + 1
                If ContadorPasajeros = 100 Then
                    ContadorPasajeros = 1
                End If
                
                parar = True
                Hoja2.Cells(j + 1, 2).Interior.ColorIndex = 3
                
                
                
            End If
            
        Wend
    End If

Next j

' Calculas las posibles ganancias.

For i = 1 To 6
    For j = 1 To 32
        If Hoja3.Cells(1 + j, i + 2) = "" Then
            
            If Hoja3.Cells(1, i + 2) = "A" Then
               PosiblesGanancias = PosiblesGanancias + Hoja3.Cells(j + 1, 11)
            End If
            If Hoja3.Cells(1, i + 2) = "B" Then
               PosiblesGanancias = PosiblesGanancias + Hoja3.Cells(j + 1, 12)
            End If
            If Hoja3.Cells(1, i + 2) = "C" Then
               PosiblesGanancias = PosiblesGanancias + Hoja3.Cells(j + 1, 13)
            End If
            If Hoja3.Cells(1, i + 2) = "D" Then
               PosiblesGanancias = PosiblesGanancias + Hoja3.Cells(j + 1, 14)
            End If
            If Hoja3.Cells(1, i + 2) = "E" Then
               PosiblesGanancias = PosiblesGanancias + Hoja3.Cells(j + 1, 15)
            End If
            If Hoja3.Cells(1, i + 2) = "F" Then
               PosiblesGanancias = PosiblesGanancias + Hoja3.Cells(j + 1, 16)
            End If
            

            
        End If
    Next j
Next i

Hoja3.Cells(6, 19) = PosiblesGanancias
    Hoja3.Range("C2:H33").Select
    Selection.Copy
    Sheets("Respuesta Heurístico").Select
    Range("B2").Select
    ActiveSheet.Paste
    Range("I3").Select
    Sheets("Avion").Select

' Hacer búsqueda local
LargoComproSilla = Range(Hoja3.Cells(1, 25), Hoja3.Cells(1, 25).End(xlDown)).Count
For k = 1 To 800
    Parar2 = False
    Cont = 0
'encontrar lugares a los que se puede cambiar.
    Hoja3.Range("v3:w100").ClearContents
    Largo = Range(Hoja3.Cells(1, 22), Hoja3.Cells(1, 22).End(xlDown)).Count
    For i = 1 To 32
        For j = 1 To 6
            If Hoja3.Cells(i + 1, j + 2) = "" Then
            
                    Largo = Range(Hoja3.Cells(1, 22), Hoja3.Cells(1, 22).End(xlDown)).Count
                    Fila = Hoja3.Cells(i + 1, j + 2).Row
                    Columna = Hoja3.Cells(i + 1, j + 2).Column
                    Cells(Largo + 1, 22) = Fila
                    Cells(Largo + 1, 23) = Columna
                    Cont = Cont + 1
            
            End If
        Next j
    Next i

' Encontrar una silla vacia
    NumSillaVacia = WorksheetFunction.RandBetween(1, Cont)
    FilSv = Hoja3.Cells(NumSillaVacia + 2, 22)
    ColSV = Hoja3.Cells(NumSillaVacia + 2, 23)
    SillaVacia = Hoja3.Cells(FilSv, ColSV)
    
' Encontrar silla para cambiar aleatoriamente
    While Parar2 = False
        RndCambioCol = WorksheetFunction.RandBetween(1, 6)
        RndCambioFil = WorksheetFunction.RandBetween(1, 32)
    
        If Hoja3.Cells(RndCambioFil + 1, RndCambioCol + 2) <> "" And Hoja3.Cells(RndCambioFil + 1, RndCambioCol + 2) <> "X" Then
           If Hoja3.Cells(RndCambioFil + 1, RndCambioCol + 2).Interior.ColorIndex <> 7 Then
                Parar2 = True
                ColCambio = RndCambioCol + 2
                FilCambio = RndCambioFil + 1
                Cambio = Hoja3.Cells(FilCambio, ColCambio)
            End If
        End If
    Wend
    
    For i = 1 To LargoComproSilla
    
    Next i
    
    
    
' Cambiar sillas
    Hoja3.Cells(FilSv, ColSV) = Hoja3.Cells(FilCambio, ColCambio)
    Hoja3.Cells(FilCambio, ColCambio) = ""

' Posibles ganancias con cambio
    PosiblesGanancias2 = 0
    For i = 1 To 6
        For j = 1 To 32
            If Hoja3.Cells(1 + j, i + 2) = "" Then
            
                
            If Hoja3.Cells(1, i + 2) = "A" Then
               PosiblesGanancias2 = PosiblesGanancias2 + Hoja3.Cells(j + 1, 11)
            End If
            If Hoja3.Cells(1, i + 2) = "B" Then
               PosiblesGanancias2 = PosiblesGanancias2 + Hoja3.Cells(j + 1, 12)
            End If
            If Hoja3.Cells(1, i + 2) = "C" Then
               PosiblesGanancias2 = PosiblesGanancias2 + Hoja3.Cells(j + 1, 13)
            End If
            If Hoja3.Cells(1, i + 2) = "D" Then
               PosiblesGanancias2 = PosiblesGanancias2 + Hoja3.Cells(j + 1, 14)
            End If
            If Hoja3.Cells(1, i + 2) = "E" Then
               PosiblesGanancias2 = PosiblesGanancias2 + Hoja3.Cells(j + 1, 15)
            End If
            If Hoja3.Cells(1, i + 2) = "F" Then
               PosiblesGanancias2 = PosiblesGanancias2 + Hoja3.Cells(j + 1, 16)
            End If
            
 
            End If
        Next j
    Next i


    If PosiblesGanancias < PosiblesGanancias2 Then
        PosiblesGanancias = PosiblesGanancias2
        Hoja3.Cells(7, 19) = PosiblesGanancias2
    Else
    Hoja3.Cells(FilSv, ColSV) = ""
    Hoja3.Cells(FilCambio, ColCambio) = Cambio
    End If

    Next k
Hoja3.Range("C2:H33").Select
    Selection.Copy
    Sheets("Respuesta Heurístico").Select
    Range("I2").Select
    ActiveSheet.Paste
    Sheets("Avion").Select
    Hoja3.Cells(8, 19) = Hoja5.Cells(6, 19)
    
    'azul
Hoja3.Range("c2:h2").Interior.ColorIndex = 8
'Verde
Hoja3.Range("c3:c12").Interior.ColorIndex = 4
'azul
Hoja3.Range("c13:c14").Interior.ColorIndex = 8
'Verde
Hoja3.Range("c15:c24").Interior.ColorIndex = 4
'amarillo
Hoja3.Range("c25:c32").Interior.ColorIndex = 6



'Verde
Hoja3.Range("e3:e12").Interior.ColorIndex = 4

'Verde
Hoja3.Range("e15:e24").Interior.ColorIndex = 4
'amarillo
Hoja3.Range("e25:e32").Interior.ColorIndex = 6

'rojo
Hoja3.Range("d3:d32").Interior.ColorIndex = 3
Hoja3.Range("g3:g32").Interior.ColorIndex = 3

'Verde
Hoja3.Range("f3:f12").Interior.ColorIndex = 4
'Verde
Hoja3.Range("f15:f24").Interior.ColorIndex = 4
'amarillo
Hoja3.Range("f25:f32").Interior.ColorIndex = 6

'Verde
Hoja3.Range("h3:h12").Interior.ColorIndex = 4
'Verde
Hoja3.Range("h15:h24").Interior.ColorIndex = 4
'amarillo
Hoja3.Range("h25:h32").Interior.ColorIndex = 6
'azul
Hoja3.Range("c13:h14").Interior.ColorIndex = 8
    
    
    
End Sub


Sub AvionReal()

Hoja5.Range("c2:h33").ClearContents
CantidadPasajeros = Range(Hoja2.Cells(1, 1), Hoja2.Cells(1, 1).End(xlDown)).Count - 1
For i = 1 To CantidadPasajeros
    Filita = Hoja2.Cells(i + 1, 11)
    Columnita = Hoja2.Cells(i + 1, 13)
    Hoja5.Cells(Filita + 1, Columnita + 2) = Hoja2.Cells(i + 1, 2)
Next i
' Ganancias Reales
GananciasReales = 0
For i = 1 To 6
    For j = 1 To 32
        If Hoja5.Cells(j + 1, i + 2) = "" Then
            If Hoja5.Cells(1, i + 2) = "A" Then
               GananciasReales = GananciasReales + Hoja5.Cells(j + 1, 11)
            End If
            
            If Hoja5.Cells(1, i + 2) = "B" Then
               GananciasReales = GananciasReales + Hoja5.Cells(j + 1, 12)
            End If
                           
            If Hoja5.Cells(1, i + 2) = "C" Then
               GananciasReales = GananciasReales + Hoja5.Cells(j + 1, 13)
            End If
            
            If Hoja5.Cells(1, i + 2) = "D" Then
               GananciasReales = GananciasReales + Hoja5.Cells(j + 1, 14)
            End If

            If Hoja5.Cells(1, i + 2) = "E" Then
               GananciasReales = GananciasReales + Hoja5.Cells(j + 1, 15)
            End If

            If Hoja5.Cells(1, i + 2) = "F" Then
               GananciasReales = GananciasReales + Hoja5.Cells(j + 1, 16)
            End If
        End If
             
        
    Next j
Next i

Hoja5.Cells(6, 19) = GananciasReales


End Sub


