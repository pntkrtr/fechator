#include "vbcompat.bi" ' Para manejo de fechas
Declare Sub muestraCalendario(ByRef mes As Integer, ByRef anno As Integer)
Declare Function rellenaPorLaIzquierda(ByVal valor As Integer, ByVal longitud As Integer, ByVal caracter As String) As String
Dim mes As Integer = Month(Now)
Dim anno As Integer = Year(Now)
muestraCalendario(mes,anno)
' *** Logica del programa
Dim entrada As String
Do While Not (entrada = "Q" Or entrada = Chr(27))
entrada = UCASE(inkey())
IF entrada = "A" Or entrada = Chr(255) + Chr(75) And anno>1970 Then anno=anno-1:muestraCalendario(mes,anno)
IF entrada = "D" Or entrada = Chr(255) + Chr(77) And anno<2099 Then anno=anno+1:muestraCalendario(mes,anno)
IF entrada = "W" Or entrada = Chr(255) + Chr(72) And mes<12 Then mes=mes+1:muestraCalendario(mes,anno)
If entrada = "S" Or entrada = Chr(255) + Chr(80) And mes>1 Then mes=mes-1:muestraCalendario(mes,anno)
Loop
Cls
Print "Calendario. Victor M. Espinosa."
Print "Garrucha a 30 de Julio de 2020"
Print "Ha finalizado el programa."
Print "Pulse una tecla para salir."
Print "Gracias por usar el programa."
Sleep
End
Sub muestraCalendario(ByRef mes As Integer, ByRef anno As Integer)
	Cls
	Dim diaSemanaPrimeroMes As Integer = WeekDay(dateserial(anno,mes,1),2)
	Dim cuenta As Integer=1
	Print "Calendario. Victor M. Espinosa."
	Print "Garrucha a 30 de Julio de 2020"
	Print
	Print "     " + MonthName(mes) + " de " + Str(anno)
	Print "---------------------------"
	Print
	Print " Lu  Ma  Mi  Ju  Vi  Sa  Do"
	Print
	Dim fila As Integer=0
	Dim diaFueraDeRango As Integer=0
	While diaFueraDeRango=0
		fila=fila+1
		Print " ";
		For columna As Integer=0 to 6
			If columna>4 Then Color 12 Else Color 11
			If fila=1 And columna<diaSemanaPrimeroMes-1 Then
				Print " .  ";
			Else
				If cuenta<=28 Then 
					Print rellenaPorLaIzquierda(cuenta, 2, " "); "  ";
					cuenta=cuenta+1
				Else
					If cuenta>28 And IsDate(Str(cuenta)+"/"+Str(mes)+"/"+Str(anno)) Then
						Print rellenaPorLaIzquierda(cuenta, 2, " "); "  ";
						cuenta=cuenta+1
						If cuenta=31 Then diaFueraDeRango=1
					Else
						Print " .  ";
						diaFueraDeRango=1
					End If
				End If
			End If
			Color 7
		Next
		Print
		Print
	Wend
	Print "---------------------------"
	Print "Use las teclas del cursor"
	Print " o A-D-S-W para moverse"
	Print "  por los calendarios"
End Sub
' Subrutina que rellena n caracteres a la izquierda de una cadena dependiendo de la longitud de la misma
Function rellenaPorLaIzquierda(ByVal valor As Integer, ByVal longitud As Integer, ByVal caracter As String) As String
	Dim caracteres As String = ""
	Dim cadena As String = Str(valor)
	For i As Integer = 0 To longitud - Len(cadena) -1
		caracteres = caracteres + caracter
	Next
	Return caracteres + str(cadena)
End Function
