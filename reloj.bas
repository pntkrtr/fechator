' Escondemos la ventana de consola
#include once "windows.bi"
Dim As HWND hwnd=GetConsoleWindow
ShowWindow(hwnd,SW_HIDE)
' Ajustamos la ventana
Screen 13
Screenres 200,200
' Definicion de Variables
Const PI As Double = 3.1415926535897932
Dim hora As String
Dim minuTo As String
Dim segundo As String
Declare Sub Circulo(ByVal x As Integer, ByVal y As Integer, ByVal radius As Integer)
' Empezamos el bucle
Do
	hora = Left(Time,2)
	minuTo = Mid(Time,4,2)
	segundo = Right(Time,2)
	Locate 1,10: Print Time
	' Dibuja la esfera
	Circulo(100,100, 80)
	' Dibuja lAs marcas minutariAs
	For n As Integer = 0 To 60
		Line ((Cos((n-15)/(pi*3.04))*78)+100,_
			(Sin((n-15)/(pi*3.04))*78)+100)-_
			((Cos((n-15)/(pi*3.04))*82)+100,_
			(Sin((n-15)/(pi*3.04))*82)+100), 100
	Next
	' Dibuja lAs marcAs horariAs
	For n As Integer = 0 To 12
		Line ((Cos((n-3)/pi*1.645)*76)+100,_
			(Sin((n-3)/pi*1.645)*76)+100)-_
			((Cos((n-3)/pi*1.645)*84)+100,_
			(Sin((n-3)/pi*1.645)*84)+100), 14
	Next
	'Dibuja el segundero
	Line (100,100)-(_
		(Cos((Val(segundo)-15)/(pi*3.04))*75)+100,_
		(Sin((Val(segundo)-15)/(pi*3.04))*75)+100), 100
	Line (100,100)-(_
		(Cos((Val(segundo)-16)/(pi*3.04))*75)+100,_
		(Sin((Val(segundo)-16)/(pi*3.04))*75)+100), 0
	' Dibuja el minutero
	Line (100,100)-(_
		(Cos((Val(minuTo)-15)/(pi*3.04))*60)+100,_
		(Sin((Val(minuTo)-15)/(pi*3.04))*60)+100), 15
	Line (100,100)-(_
		(Cos((Val(minuTo)-16)/(pi*3.04))*60)+100,_
		(Sin((Val(minuTo)-16)/(pi*3.04))*60)+100), 0
	' Dibuja la aguja horaria
	Line (100,100)-(_
		(Cos((Val(hora)-3)/pi*1.645)*40)+100,_
		(Sin((Val(hora)-3)/pi*1.645)*40)+100), 14
	Line (100,100)-(_
		(Cos((Val(hora)-4)/pi*1.645)*40)+100,_
		(Sin((Val(hora)-4)/pi*1.645)*40)+100), 0
' Solo sale si pulsamos espacio
Loop While Not Inkey=" "
ShowWindow(hwnd,SW_SHOW)
End
Sub Circulo(ByVal x As Integer, ByVal y As Integer, ByVal radius As Integer)
    For i As double = 0 To PI * 2 Step 0.005
        Var x2 = Cos(i), y2 = Sin(i)
        Pset ((x2 * radius) + x, (y2 * radius) + y)        
    Next
End Sub
