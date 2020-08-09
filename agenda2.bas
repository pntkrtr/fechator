#include "vbcompat.bi" ' Para manejo de fechas
/' Pendiente:
	- Al insertar registros validar el campo hora con isDate()
	- Modificar registros, pasar a pasivo, marchar como hecho, ...
	- Eliminar registros.
	- Controlar la lógica de la fecha fin (no puede ser inferior a la inicial, ...)
	- Opción para forzar el guardado de los datos (puede que no sea necesario)
	- Al salir, hacer copia de seguridad y guardar los datos en el archivo (puede que no sea necesario)
	- Alarmas de eventos en el momento en el que suceden
	- Posible: Interfaz gráfica o modo texto calendario.
'/
' Declaraciones
Declare Function seguridad As String
Declare Function numElemBD As Integer
Declare Sub cargaDatos
Declare Function rellenaPorLaIzquierda(ByVal valor As Integer, ByVal longitud As Integer, ByVal caracter As String) As String
Declare Sub refrescaPantalla(ByVal id As Integer, _
	ByVal activo As Integer, ByVal completado As Integer, _
	ByVal fechaInicio As String, ByVal horaInicio As String, _
	ByVal fechaFin As String, ByVal horaFin As String, _
	ByVal info As String, ByVal etiquetas As String, _
	ByVal personas As String, ByVal lugares As String)
Declare Sub muestraCabecera(ByVal texto As String)
Declare Sub muestraDatos(ByVal id As Integer, _
	ByVal activo As Integer, ByVal completado As Integer, _
	ByVal fechaInicio As String, ByVal horaInicio As String, _
	ByVal fechaFin As String, ByVal horaFin As String, _
	ByVal info As String, ByVal etiquetas As String, _
	ByVal personas As String, ByVal lugares As String)
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
	personas As String
	lugares As String
End Type
' Variables del programa
Dim Shared ver As String
Dim Shared usr As String
Dim Shared nomArchivo As String
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
IF entrada = "A" And posicion>Lbound(ev)+1 Then 
	posicion=posicion-1
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).completado, _
		ev(posicion).fechaInicio, ev(posicion).horaInicio, _
		ev(posicion).fechaFin, ev(posicion).horaFin, _
		ev(posicion).info, ev(posicion).etiquetas, _
		ev(posicion).personas, ev(posicion).lugares)
End If
IF entrada = "D" And posicion<Ubound(ev) Then 
	posicion=posicion+1
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).completado, _
		ev(posicion).fechaInicio, ev(posicion).horaInicio, _
		ev(posicion).fechaFin, ev(posicion).horaFin, _
		ev(posicion).info, ev(posicion).etiquetas, _
		ev(posicion).personas, ev(posicion).lugares)
End If
IF entrada = "U" Then 
	usr=seguridad
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).completado, _
		ev(posicion).fechaInicio, ev(posicion).horaInicio, _
		ev(posicion).fechaFin, ev(posicion).horaFin, _
		ev(posicion).info, ev(posicion).etiquetas, _
		ev(posicion).personas, ev(posicion).lugares)
End If
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
		nomArchivo = "agenda-" + usr + "-2.dat"
		' Si no existe el archivo, lo creamos
		If Not Fileexists(nomArchivo) then
			Open nomArchivo For Output As #1: Close #1
		End If
		' Calculamos en número de elementos del array
		Dim maximo As Integer = numElemBD
		' Dimensionamos el array para la carga de datos y lo compartimos
		Redim ev(maximo)
		' Cargamos los datos en el array
		cargaDatos
		' Ordenamos el array de datos cargado
		ordenaArray
		' Mostramos los datos del primer registro
		posicion = 1
		refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).completado, _
			ev(posicion).fechaInicio, ev(posicion).horaInicio, _
			ev(posicion).fechaFin, ev(posicion).horaFin, _
			ev(posicion).info, ev(posicion).etiquetas, _
			ev(posicion).personas, ev(posicion).lugares)
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
	Dim maximo As Integer = 0
	While Not Eof(1)
		Dim As Integer activoTmp, completadoTmp
		Dim As String fechaInicioTmp, horaInicioTmp, fechaFinTmp, horaFinTmp
		Dim As String infoTmp, etiquetasTmp, personasTmp, lugaresTmp
		Input #1, activoTmp, completadoTmp, _
			fechaInicioTmp, horaInicioTmp, _
			fechaFinTmp, horaFinTmp, _
			infoTmp, etiquetasTmp, _
			personasTmp, lugaresTmp
		If activoTmp = 1 Or muestraPasivos = 1 Then maximo = maximo + 1
	Wend
	Close #1
	Return maximo
End Function
' Subrutina que carga los datos en el array
Sub cargaDatos
	Dim posicion As Integer = 0
	Open nomArchivo For Input As #1
	Dim As Integer activoTmp, completadoTmp
	Dim As String fechaInicioTmp, horaInicioTmp, fechaFinTmp, horaFinTmp
	Dim As String infoTmp, etiquetasTmp, personasTmp, lugaresTmp
	While Not Eof(1)
		Input #1, activoTmp, completadoTmp, _
			fechaInicioTmp, horaInicioTmp, _
			fechaFinTmp, horaFinTmp, _
			infoTmp, etiquetasTmp, _
			personasTmp, lugaresTmp
		If activoTmp = 1 Or muestraPasivos = 1 Then
			ev(posicion).activo = activoTmp
			ev(posicion).completado = completadoTmp
			ev(posicion).fechaInicio = fechaInicioTmp
			ev(posicion).horaInicio = horaInicioTmp
			ev(posicion).fechaFin = fechaFinTmp
			ev(posicion).horaFin = horaFinTmp
			ev(posicion).info = infoTmp
			ev(posicion).etiquetas = etiquetasTmp
			ev(posicion).personas = personasTmp
			ev(posicion).lugares = lugaresTmp
			posicion = posicion + 1
		End If
	Wend
	Close #1
End Sub
' Funcion que llama a las subrutinas que componen la pantalla
Sub refrescaPantalla(ByVal id As Integer, _
	ByVal activo As Integer, ByVal completado As Integer, _
	ByVal fechaInicio As String, ByVal horaInicio As String, _
	ByVal fechaFin As String, ByVal horaFin As String, _
	ByVal info As String, ByVal etiquetas As String, _
	ByVal personas As String, ByVal lugares As String)
	muestraCabecera("Calendario y agenda: Recorrido individual de registros")
	muestraDatos id, activo, completado, _
			fechaInicio, horaInicio, _
			fechaFin, horaFin, _
			info, etiquetas, _
			personas, lugares
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas y llama al resto de las subrutinas que componen la pantalla
Sub muestraListado
	muestraCabecera("Calendario y agenda: Listado de registros")
	For id As Integer = Lbound(ev)+1 To Ubound(ev)
		muestraDatos id, ev(id).activo, ev(id).completado, _
			ev(id).fechaInicio, ev(id).horaInicio, _
			ev(id).fechaFin, ev(id).horaFin, _
			ev(id).info, ev(id).etiquetas, _
			ev(id).personas, ev(id).lugares
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas de hoy y llama al resto de las subrutinas que componen la pantalla
Sub muestraListadoHoy
	muestraCabecera("Calendario y agenda: Listado de hoy")
	For id As Integer = Lbound(ev)+1 To Ubound(ev)
		' Comprobamos la correspondencia con el día de hoy
		If (panio(ev(id).fechaInicio)=Str(Year(Now)) Or panio(ev(id).fechaInicio)="0000") _
			And (pmes(ev(id).fechaInicio)=Str(rellenaPorLaIzquierda(Month(Now),2,"0")) Or pmes(ev(id).fechaInicio)="00") _
			And	pdia(ev(id).fechaInicio)=Str(rellenaPorLaIzquierda(Day(Now),2,"0")) Then
				muestraDatos id, ev(id).activo, ev(id).completado, _
					ev(id).fechaInicio, ev(id).horaInicio, _
					ev(id).fechaFin, ev(id).horaFin, _
					ev(id).info, ev(id).etiquetas, _
					ev(id).personas, ev(id).lugares
		End If
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas de esta semana y llama al resto de las subrutinas que componen la pantalla
Sub muestraListadoEstaSemana
	muestraCabecera("Calendario y agenda: Listado de esta semana")
	For id As Integer = Lbound(ev)+1 To Ubound(ev)
		' Comprobamos la correspondencia con esta semana (Falta incluir eventos repetitivos)
		If DateValue(ev(id).fechaInicio)>=Now And DateValue(ev(id).fechaInicio)<=Now+7 Then
				muestraDatos id, ev(id).activo, ev(id).completado, _
					ev(id).fechaInicio, ev(id).horaInicio, _
					ev(id).fechaFin, ev(id).horaFin, _
					ev(id).info, ev(id).etiquetas, _
					ev(id).personas, ev(id).lugares
		End If
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas de este mes y llama al resto de las subrutinas que componen la pantalla
Sub muestraListadoEsteMes
	muestraCabecera("Calendario y agenda: Listado de este mes")
	For id As Integer = Lbound(ev)+1 To Ubound(ev)
		' Comprobamos la correspondencia con este mes
		If (panio(ev(id).fechaInicio)=Str(Year(Now)) Or panio(ev(id).fechaInicio)="0000") _
			And (pmes(ev(id).fechaInicio)=Str(rellenaPorLaIzquierda(Month(Now),2,"0")) Or pmes(ev(id).fechaInicio)="00") Then 
			muestraDatos id, ev(id).activo, ev(id).completado, _
				ev(id).fechaInicio, ev(id).horaInicio, _
				ev(id).fechaFin, ev(id).horaFin, _
				ev(id).info, ev(id).etiquetas, _
				ev(id).personas, ev(id).lugares
		End If
	Next
	muestraInstrucciones
	muestraGraficos
End Sub
' Funcion que saca el listado de citas de este anio y llama al resto de las subrutinas que componen la pantalla
Sub muestraListadoEsteAnio
	muestraCabecera("Calendario y agenda: Listado de este anio")
	For id As Integer = Lbound(ev)+1 To Ubound(ev)
		' Comprobamos la correspondencia con este anio
		If (panio(ev(id).fechaInicio)=Str(Year(Now)) Or panio(ev(id).fechaInicio)="0000") Then
			muestraDatos id, ev(id).activo, ev(id).completado, _
				ev(id).fechaInicio, ev(id).horaInicio, _
				ev(id).fechaFin, ev(id).horaFin, _
				ev(id).info, ev(id).etiquetas, _
				ev(id).personas, ev(id).lugares
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
	For id As Integer = Lbound(ev)+1 To Ubound(ev)
		' Comprobamos la correspondencia con esta semana (Falta incluir casos repetitivos)
		If DateValue(ev(id).fechaInicio)>=DateValue(fechaInicio_inicio) And DateValue(ev(id).fechaInicio)<=DateValue(fechaInicio_fin) Then
			muestraDatos id, ev(id).activo, ev(id).completado, _
				ev(id).fechaInicio, ev(id).horaInicio, _
				ev(id).fechaFin, ev(id).horaFin, _
				ev(id).info, ev(id).etiquetas, _
				ev(id).personas, ev(id).lugares
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
Sub muestraDatos(ByVal id As Integer, _
	ByVal activo As Integer, ByVal completado As Integer, _
	ByVal fechaInicio As String, ByVal horaInicio As String, _
	ByVal fechaFin As String, ByVal horaFin As String, _
	ByVal info As String, ByVal etiquetas As String, _
	ByVal personas As String, ByVal lugares As String)
	If activo=1 Or muestraPasivos=1 Then 
		Color 6:Print rellenaPorLaIzquierda(id, 3, " ") ;
		Color 7:Print " | " ; activo ; "| " ; completado ; "| " ; 
		if activo=1 then Color 10 Else Color 14
		Print fechaInicio ; " - " ; horaInicio ; 
		Color 7:Print " | " ; 
		if activo=1 then Color 10 Else Color 14
		Print fechaFin ; " - " ; horaFin ; 
		Color 7:Print " | " ; 
		Color 31: Print info
		Color 7:Print " | " ; 
		Color 31: Print etiquetas ;
		Color 7:Print " | " ; 
		Color 31: Print personas ;
		Color 7:Print " | " ; 
		Color 31: Print lugares
	End If
End Sub
' Subrutina que muestra instrucciones
Sub muestraInstrucciones()
	Locate 59,60:Print "Usuario conectado: " ; usr
	Locate 60,60:Print "Numero de registros cargados:" ; Ubound(ev)
	Locate 61,60:Print "Posicion actual:" ; posicion
	Locate 62,60:Print "Q o ESC para terminar el programa"
	Locate 63,60:Print "U para cerrar sesion"
	Locate 64,60:Print "A y D para desplazarse por los registros"
	Locate 65,60:Print "L para un listado de todas las citas"
	Locate 66,60:Print "T para ver las citas de hoy"
	Locate 67,60:Print "W para ver las citas de hoy a una semana"
	Locate 68,60:Print "M para ver las citas de este mes"
	Locate 69,60:Print "Y para ver las citas de este anio"
	Locate 70,60:Print "C para ver las citas entre dos fechas"
	Locate 71,60:Print "O para mostrar/ocultar pasivos: " ; 
	If muestraPasivos=1 Then Print "ON" Else Print "OFF"
	Locate 72,60:Print "N para registrar nuevo evento"
	Locate 73,60:Print "H para informacion completa."
End Sub
' Subrutnia que muestra graficos
Sub muestraGraficos
	Line (450,450)-(800,595),15,b
End Sub
Sub switchMuestraPasivos
	if muestraPasivos = 0 Then muestraPasivos = 1 Else muestraPasivos = 0
	' Dimensionamos el array con el número de elementos
	Redim ev(numElemBD)
	' Cargamos los datos en el array
	cargaDatos
	' Ordenamos el array
	ordenaArray
	' Mostramos los datos del primer registro
	posicion = 1
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).completado, _
			ev(posicion).fechaInicio, ev(posicion).horaInicio, _
			ev(posicion).fechaFin, ev(posicion).horaFin, _
			ev(posicion).info, ev(posicion).etiquetas, _
			ev(posicion).personas, ev(posicion).lugares)
End Sub
' Registrar nuevo evento
Sub nuevoEvento
	Cls
	Print "Registro de nuevo evento (En construccion)"
	' Recoge datos
	Dim miEvento As evento
	' --- Hay que ir validando los valores introducidos
	Do
		Input "Activo (0-1)" ; miEvento.activo
		If miEvento.activo < 0 Or miEvento.activo > 1 Then Print "Debe introducir un valor dentro del rango"
	Loop While miEvento.activo < 0 Or miEvento.activo > 1
	Do 
		Input "Completado (0-1)" ; miEvento.completado
	If miEvento.completado < 0 Or miEvento.completado > 1 Then Print "Debe introducir un valor dentro del rango"
	Loop While miEvento.activo < 0 Or miEvento.activo > 1
	Do
		Input "Fecha de incicio (dd/mm/aaaa)" ; miEvento.fechaInicio
		If Not IsDate(miEvento.fechaInicio) Then Print "Debe introducir una fecha correcta"
	Loop While Not IsDate(miEvento.fechaInicio)
	' --- Comprobar si la hora introducida es válida
	Input "Hora de inicio (hh:mm)" ; miEvento.horaInicio
	Do 
		Input "Fecha de fin (dd/mm/aaaa)"; miEvento.fechaFin
		If Not IsDate(miEvento.fechaFin) Then Print "Debe introducir una fecha correcta"
	Loop While Not IsDate(miEvento.fechaFin)
	' --- Comprobar si la hora introducida es válida
	Input "Hora de fin (hh:mm)" ; miEvento.horaFin
	Do 
		Input "Información: " , miEvento.info
		If miEvento.info = "" Then Print "Debe introducir algún texto para el evento"
	Loop While miEvento.info = ""
	Input "Etiquetas (etiqueta1|etiqueta2|etiqueta3|...: " , miEvento.etiquetas
	Input "Personas (persona1|persona2|persona3|...): " , miEvento.personas
	Input "Lugares (lugar1|lugar2|lugar3|...): " , miEvento.lugares
	' Actualiza array
	Dim eventoTemporal(Ubound(ev)+1) As evento
	For id As Integer = Lbound(ev) To Ubound(ev)
		eventoTemporal(id) = ev(id)
	Next
	eventoTemporal(Ubound(ev)+1) = miEvento
	Redim ev(Ubound(eventoTemporal))
	For id As Integer = Lbound(ev) To Ubound(ev)
		ev(id) = eventoTemporal(id)
	Next
	' Renombramos el archivo antiguo como backup
	Dim nuevoNombreArchivo As String = nomArchivo + "." + Str(Now)
	Dim result As Integer = Name ( nomArchivo , nuevoNombreArchivo )
	If result <> 0  Then 
		Print  "Fallo al cambiar el nombre del archivo" 
		Sleep()
	End If
	' Creamos un nuevo archivo
	If Not Fileexists(nomArchivo) then
		Open nomArchivo For Output As #1
	Else
		Print "No se ha podido crear el archivo porque ya existe"
		Close #1
		Sleep()
	End If
	' Escribe en el archivo
	For id As Integer = 1 To Ubound(ev)
		Print #1, Str(ev(id).activo) + "," + Str(ev(id).completado) + "," _
			+ ev(id).fechaInicio + "," + ev(id).horaInicio + "," _
			+ ev(id).fechaFin + "," + ev(id).horaFin + "," + Chr(34) + ev(id).info + Chr(34) _
			+ "," + Chr(34) + ev(id).etiquetas + Chr(34) + "," + Chr(34) + ev(id).personas + Chr(34) _
			+ "," + Chr(34) + ev(id).lugares + Chr(34)
	Next
	Close #1
	' Ordena el array
	 ordenaArray
	' Recarga Vista
	Sleep
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).completado, _
			ev(posicion).fechaInicio, ev(posicion).horaInicio, _
			ev(posicion).fechaFin, ev(posicion).horaFin, _
			ev(posicion).info, ev(posicion).etiquetas, _
			ev(posicion).personas, ev(posicion).lugares)
End Sub
' Mostramos información de interés
Sub Ayuda
	Cls
	Print "Calendario y Agenda. Alfa concept preview. " + ver + ". Victor M. Espinosa"
	Print "Informacion de interes"
	Sleep
	refrescaPantalla(posicion, ev(posicion).activo, ev(posicion).completado, _
			ev(posicion).fechaInicio, ev(posicion).horaInicio, _
			ev(posicion).fechaFin, ev(posicion).horaFin, _
			ev(posicion).info, ev(posicion).etiquetas, _
			ev(posicion).personas, ev(posicion).lugares)
End Sub
' Ordenamos el array por fecha
Sub ordenaArray
	' Algoritmo de ordenación
	For n As Integer = Lbound(ev) To Ubound(ev)-1
		For i As Integer = Lbound(ev) To Ubound(ev)-1
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
