#include "vbcompat.bi" ' Para manejo de fechas
/' Pendiente:
	- Al entrar, si el archivo está vacío: comprobar qué sucede
	- Insertar registro, validar fecha isDate() + ordenar + ¿guardar datos?
	- Al salir, hacer copia de seguridad y guardar los datos en el archivo 
	- Alarmas de eventos cuando en el momento en el que suceden
	- Posible: Distintos Usuarios
	- Posible: Múltiples bases de datos por usuario
	- Posible: Interfaz gráfica o modo texto calendario.
'/
' Declaraciones
Declare Function seguridad As String
Declare Function numElemBD As Integer
Declare Sub cargaDatos
Declare Function rellenaPorLaIzquierda(ByVal valor As Integer, ByVal longitud As Integer, ByVal caracter As String) As String
Declare Sub refrescaPantalla(ByVal id As Integer, ByVal activo As Integer, ByVal fecha As String, ByVal hora As String, ByVal info As String)
Declare Sub muestraCabecera(ByVal texto As String)
Declare Sub muestraDatos(ByVal id As Integer, ByVal activo As Integer, ByVal fecha As String, ByVal hora As String, ByVal info As String)
Declare Sub muestraListado
Declare Sub muestraListadoHoy
Declare Sub muestraListadoEstaSemana
Declare Sub muestraListadoEsteMes
Declare Sub muestraListadoEsteAnio
Declare Sub muestraListadoEntreFechas
Declare Sub muestraInstrucciones
Declare Sub muestraGraficos
Declare Sub listaCitas
Declare Sub switchMuestraPasivos
Declare Sub nuevoEvento
Declare Sub ayuda
Declare Sub ordenaArray
Declare Function cadenaNormalizadaFechaHora(ByVal fecha As String, ByVal hora As String) As String
Declare Function panio(ByVal fecha As String) As String
Declare Function pmes(ByVal fecha As String) As String
Declare Function pdia(ByVal fecha As String) As String
Declare Function phora(ByVal hora As String) As String
Declare Function pminutos(ByVal hora As String) As String
' Tipos
Type evento
	activo As Integer
	completado As Integer
	fechaInicio As String
	horaInicio As String
	fechaFin As String
	horaFin As String
	info As String
	etiquetas As String
End Type
' Variables del programa
Dim Shared ver As String
Dim Shared usr As String
Dim Shared nomArchivo As String
Dim Shared maximo As Integer
Dim Shared posicion As Integer
Dim Shared muestraPasivos As Integer = 1
Redim Shared ev() As evento
' Tamaño y resolucion de la ventana
ScreenRes 800,600
Window Screen (0,0)-(800,600)
' Seguridad
usr = seguridad
' *** Logica del programa
ver = "20200712"
Dim entrada As String
Do While Not (entrada = "Q" Or entrada = Chr(27))
entrada = UCASE(inkey())
IF entrada = "A" And posicion>0 Then posicion=posicion-1:refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).fechaInicio, ev(posicion).horaInicio, ev(posicion).info)
IF entrada = "D" And posicion<maximo-1 Then posicion=posicion+1:refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).fechaInicio, ev(posicion).horaInicio, ev(posicion).info)
IF entrada = "U" Then usr=seguridad:refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).fechaInicio, ev(posicion).horaInicio, ev(posicion).info)
If entrada = "L" Then muestraListado
If entrada = "T" Then muestraListadoHoy
If entrada = "W" Then muestraListadoEstaSemana
If entrada = "M" Then muestraListadoEsteMes
If entrada = "Y" Then muestraListadoEsteAnio
If entrada = "C" Then muestraListadoEntreFechas
If entrada = "O" Then switchMuestraPasivos
If entrada = "N" Then nuevoEvento
If entrada = "H" Then ayuda
Loop
Cls
Print "Calendario y Agenda. Alfa concept preview. " + ver + ". Victor M. Espinosa"
Print "Ha finalizado el programa. Pulse una tecla para salir."
Print "Gracias por usar el programa."
Sleep
End
' *** Subrutinas y funciones ***
' Funcion que comprueba las credenciales y devuelve el usuario si es valido
Function seguridad() As String
	Dim As String pwd
	Cls
	Input "Usuario:", usr
	Input "Clave:", pwd
	If pwd="123456" And Len(usr)>0 Then
		usr=Lcase(usr)
		' Generamos un nombre de archivo
		nomArchivo = "agenda-" + usr + ".dat"
		' Si no existe el archivo, lo creamos
		If Not Fileexists(nomArchivo) then
			Open nomArchivo For Output As #1: Close #1
		End If
		' Calculamos en número de elementos del array
		maximo = numElemBD
		' Dimensionamos el array para la carga de datos y lo compartimos
		Redim ev(maximo)
		' Cargamos los datos en el array
		cargaDatos
		' Ordenamos el array de datos cargado
		ordenaArray
		' Mostramos los datos del primer registro
		posicion = 0
		refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).fechaInicio, ev(posicion).horaInicio, ev(posicion).info)
		Return usr
	Else
		Cls 
		Print "Calendario y Agenda. Alfa concept preview. " + ver + ". Víctor M. Espinosa)"
		Print "Usuario o clave incorrectos. Finalizando la aplicacion."
		Print "Gracias por intentar usar el programa."
		End
	End If
End Function
' Funcion que devuelve el numero de elementos de la BD
Function numElemBD As Integer
	Open nomArchivo For Input As #1
	Dim linea As String
	Dim maximo As Integer = 0
	While Not Eof(1)
		Dim activoTmp As Integer
		Dim As String fechaTmp, horaTmp, infoTmp
		Input #1, activoTmp, fechaTmp, horaTmp, infoTmp
		If activoTmp = 1 Or muestraPasivos = 1 Then maximo = maximo + 1
	Wend
	Close #1
	Return maximo
End Function
' Subrutina que carga los datos en el array
Sub cargaDatos
	Dim posicion As Integer = 0
	Open nomArchivo For Input As #1
	While Not Eof(1)
		Dim activoTmp As Integer
		Dim As String fechaTmp, horaTmp, infoTmp
		Input #1, activoTmp, fechaTmp, horaTmp, infoTmp
			ev(posicion).activo=activoTmp
			ev(posicion).fechaInicio=fechaTmp
			ev(posicion).horaInicio=horaTmp
			ev(posicion).info=infoTmp
			posicion = posicion + 1
	Wend
	Close #1
End Sub
' Funcion que llama a las subrutinas que componen la pantalla
Sub refrescaPantalla(ByVal id As Integer, ByVal activo As Integer, ByVal fecha As String, ByVal hora As String, ByVal info As String)
	muestraCabecera("Calendario y agenda: Recorrido individual de registros")
	muestraDatos id, activo, fecha, hora, info
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas y llama al resto de las subrutinas que componen la pantalla
Sub muestraListado
	muestraCabecera("Calendario y agenda: Listado de registros")
	For id As Integer = 0 To maximo-1
		muestraDatos id, ev(id).activo, ev(id).fechaInicio, ev(id).horaInicio, ev(id).info
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas de hoy y llama al resto de las subrutinas que componen la pantalla
Sub muestraListadoHoy
	muestraCabecera("Calendario y agenda: Listado de hoy")
	For id As Integer = 0 To maximo-1
		' Comprobamos la correspondencia con el día de hoy
		If (panio(ev(id).fechaInicio)=Str(Year(Now)) Or panio(ev(id).fechaInicio)="0000") _
			And (pmes(ev(id).fechaInicio)=Str(rellenaPorLaIzquierda(Month(Now),2,"0")) Or pmes(ev(id).fechaInicio)="00") _
			And	pdia(ev(id).fechaInicio)=Str(rellenaPorLaIzquierda(Day(Now),2,"0")) Then
				muestraDatos id, ev(id).activo, ev(id).fechaInicio, ev(id).horaInicio, ev(id).info
		End If
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas de esta semana y llama al resto de las subrutinas que componen la pantalla
Sub muestraListadoEstaSemana
	muestraCabecera("Calendario y agenda: Listado de esta semana")
	For id As Integer = 0 To maximo-1
		' Comprobamos la correspondencia con esta semana (Falta incluir eventos repetitivos)
		If DateValue(ev(id).fechaInicio)>=Now And DateValue(ev(id).fechaInicio)<=Now+7 Then
				muestraDatos id, ev(id).activo, ev(id).fechaInicio, ev(id).horaInicio, ev(id).info
		End If
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas de este mes y llama al resto de las subrutinas que componen la pantalla
Sub muestraListadoEsteMes
	muestraCabecera("Calendario y agenda: Listado de este mes")
	For id As Integer = 0 To maximo-1
		' Comprobamos la correspondencia con este mes
		If (panio(ev(id).fechaInicio)=Str(Year(Now)) Or panio(ev(id).fechaInicio)="0000") _
			And (pmes(ev(id).fechaInicio)=Str(rellenaPorLaIzquierda(Month(Now),2,"0")) Or pmes(ev(id).fechaInicio)="00") Then 
			muestraDatos id, ev(id).activo, ev(id).fechaInicio, ev(id).horaInicio, ev(id).info
		End If
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas de este anio y llama al resto de las subrutinas que componen la pantalla
Sub muestraListadoEsteAnio
	muestraCabecera("Calendario y agenda: Listado de este anio")
	For id As Integer = 0 To maximo-1
		' Comprobamos la correspondencia con este anio
		If (panio(ev(id).fechaInicio)=Str(Year(Now)) Or panio(ev(id).fechaInicio)="0000") Then
			muestraDatos id, ev(id).activo, ev(id).fechaInicio, ev(id).horaInicio, ev(id).info
		End If
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas entre fechas
Sub muestraListadoEntreFechas
	muestraCabecera("Calendario y agenda: Listado entre dos fechas")
	Dim As String fechaInicio_inicio, fechaInicio_fin
	Print "Introduzca las fechas validas de inicio y fin para la busqueda."
	While isDate(fechaInicio_inicio)=0
		Input "Fecha inicio (dd/mm/yyyy): ", fechaInicio_inicio
		If isDate(fechaInicio_inicio)=0 Then Print "Fecha incorrecta. Escriba una fecha valida."
	Wend
	While isDate(fechaInicio_fin)=0
		Input "Fecha fin    (dd/mm/yyyy): ", fechaInicio_fin
		If isDate(fechaInicio_fin)=0 Then Print "Fecha incorrecta. Escriba una fecha valida."
	Wend
	For id As Integer = 0 To maximo-1
		' Comprobamos la correspondencia con esta semana (Falta incluir casos repetitivos)
		If DateValue(ev(id).fechaInicio)>=DateValue(fechaInicio_inicio) And DateValue(ev(id).fechaInicio)<=DateValue(fechaInicio_fin) Then
				muestraDatos id, ev(id).activo, ev(id).fechaInicio, ev(id).horaInicio, ev(id).info
		End If
	Next
	muestraInstrucciones
	muestraGraficos
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
' Subrutina que muestra la cabecera
Sub muestraCabecera(ByVal texto As String)
	Cls
	Print texto
	For i As Integer = 0 To Len(texto)-1
		Print "-";
	Next
	Print
End Sub
' Subrutina que muestra los datos
Sub muestraDatos(ByVal id As Integer, ByVal activo As Integer, ByVal fecha As String, ByVal hora As String, ByVal info As String)
	If activo=1 Or muestraPasivos=1 Then 
		Color 6:Print rellenaPorLaIzquierda(id, 3, " ") ; 
		Color 7:Print " | " ; activo ; "| " ; 
		if activo=1 then Color 10 Else Color 14
		Print fecha ; " - " ; hora ; 
		Color 7:Print " | " ; 
		Color 31: Print info
	End If
End Sub
' Subrutina que muestra instrucciones
Sub muestraInstrucciones()
	Locate 3,60:Print "Usuario conectado: " ; usr
	Locate 4,60:Print "Numero de registros cargados:" ; maximo
	Locate 5,60:Print "Posicion actual:" ; posicion
	Locate 6,60:Print "Q o ESC para terminar el programa"
	Locate 7,60:Print "U para cerrar sesion"
	Locate 8,60:Print "A y D para desplazarse por los registros"
	Locate 9,60:Print "L para un listado de todas las citas"
	Locate 10,60:Print "T para ver las citas de hoy"
	Locate 11,60:Print "W para ver las citas de hoy a una semana"
	Locate 12,60:Print "M para ver las citas de este mes"
	Locate 13,60:Print "Y para ver las citas de este anio"
	Locate 14,60:Print "C para ver las citas entre dos fechas"
	Locate 15,60:Print "O para mostrar/ocultar pasivos: " ; 
	If muestraPasivos=1 Then Print "ON" Else Print "OFF"
	Locate 16,60:Print "N para registrar nuevo evento"
	Locate 17,60:Print "H para informacion completa."
End Sub
' Subrutnia que muestra graficos
Sub muestraGraficos
	Line (450,5)-(800,140),15,b
End Sub
Sub switchMuestraPasivos
	if muestraPasivos = 0 Then muestraPasivos = 1 Else muestraPasivos = 0
	' Calculamos en número de elementos del array
	maximo = numElemBD
	' Dimensionamos el array para la carga de datos y lo compartimos
	Redim ev(maximo)
	' Cargamos los datos en el array
	cargaDatos
	' Mostramos los datos del primer registro
	posicion = 0
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).fechaInicio, ev(posicion).horaInicio, ev(posicion).info)
End Sub
' Registrar nuevo evento
Sub nuevoEvento
	Cls
	Print "Registro de nuevo evento (En construccion)"
	' Recoge datos
	' Escribe en el archivo
	' Actualiza array
	' Ordena el array
	' Recarga Vista
	Sleep
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).fechaInicio, ev(posicion).horaInicio, ev(posicion).info)
End Sub
' Mostramos información de interés
Sub Ayuda
	Cls
	Print "Calendario y Agenda. Alfa concept preview. " + ver + ". Victor M. Espinosa"
	Print "Informacion de interes"
	Sleep
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).fechaInicio, ev(posicion).horaInicio, ev(posicion).info)
End Sub
' Ordenamos el array por fecha
Sub ordenaArray
	' Algoritmo de ordenación
	For n As Integer = Lbound(ev) To Ubound(ev)
		For i As Integer = Lbound(ev) To Ubound(ev)-2
			If Val(cadenaNormalizadaFechaHora(ev(i).fechaInicio,ev(i).horaInicio)) > _ 
				Val(cadenaNormalizadaFechaHora(ev(i+1).fechaInicio,ev(i+1).horaInicio)) Then Swap ev(i), ev(i+1)
		Next
	Next
End Sub
' Devolvemos una cadena con la fecha y la hora en formato AAAAMMDDHHMM
Function cadenaNormalizadaFechaHora(ByVal fecha As String, ByVal hora As String) As String
	Return panio(fecha) + pmes(fecha) + pdia(fecha) + phora(hora) + pminutos(hora)
End Function
' Si nos dan una fecha de la agenda devolvemos el anio
Function panio(ByVal fecha As String) As String
	Return Right(fecha,4)
End Function
' Si nos dan una fecha de la agenda devolvemos el mes
Function pmes(ByVal fecha As String) As String
	Return Mid(fecha,4,2)
End Function
' Si nos dan una fecha de la agenda devolvemos el dia
Function pdia(ByVal fecha As String) As String
	Return Left(fecha,2)
End Function
' Si nos dan una hora de la agenda devolvemos la parte horaria
Function phora(ByVal hora As String) As String
	Return Left(hora,2)
End Function
' Si nos dan una hora de la agenda devolvemos la parte de los minutos
Function pminutos(ByVal hora As String) As String
	Return Right(hora,2)
End Function
