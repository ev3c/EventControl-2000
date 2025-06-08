Attribute VB_Name = "Declaraciones"
' Para que no pida contrase�a si no autoarranque
Public bolPrimerArranque As Boolean

' Database y recorsets de EventControl.mdb
Public gdb As Database
Public grsPrograma As Recordset
Public grsEvento As Recordset

' Database y recorsets de Historial
Public gdbHistorial As Database
Public grsHistorial As Recordset

' Array para guardar las url del dia
Public gaURL() As String
Public gaHistorial() As String


Public Type ecOnOff
  On As Date        'Hora de arranque
  Off As Date       'Hora de parada
End Type

Public gaWin() As ecOnOff  'Evento Windows
Public gaScr() As ecOnOff  'Evento Protector Pantalla
Public gaMdm() As ecOnOff  'Evento M�dem
Public gaTmr() As ecOnOff  'Evento Cronometro
Public gaPrg() As ecOnOff  'Evento Programa

Public giWin As Integer   'Indice para evento Windows
Public giScr As Integer   '�ndice para evento Pantalla
Public giMdm As Integer   '�ndice para evento M�dem
Public giTmr As Integer   '�ndice para evento Cronometro
Public gnPrg As Integer   '�ndice para n�mero Programa
Public giPrg() As Integer 'Indice para evento Programa

Public gtWin As ecOnOff   'Temp �ltimo evento Windows
Public gtScr As ecOnOff   'Temp �ltimo evento Screen
Public gtMdm As ecOnOff   'Temp �ltimo evento Modem
Public gtTmr As ecOnOff   'Temp �ltimo evento Cronometro
Public gtPrg() As ecOnOff 'Temp �ltimos eventos Programas

Public gdPrgDia() As Date  'Horas programas dia
Public gdPrgMes() As Date  'Horas programas mes
Public gdPrgA�o() As Date  'Horas programas a�o

' Fecha en que arranca el Programa EventControl
Public gdFechaOn As Date

'Para control de programas introducidos por el usuario
Public gaPrograma(0 To 99, 0 To 2) As String

Public gstrIconMsg As String      'Mensaje del Icono
Public gstrPrograma As String     'Nombre y Versi�n de EventControl
Public gstrContrase�a As String   'Contrase�a
Public gblnContrase�a As Boolean  'Contrase�a correcta
Public gstrFormatoFecha As String 'Formato de la fecha
Public gblnCronoOn As Boolean     'Cronometro On/Off

'Public garrFichero(99) As String  'Para guardar los programas activos

Public Const EC_ESPA�OL = 0
Public Const EC_ENGLISH = 1
Public Const EC_CATALA = 2

Public EC_SHAREWARE As Boolean  'True si coincide Encriptado
Public gstrNumeroHD As String   'Para guardar Numero HD
Public gstrNumeroRegistro As String  'Para guardar Num Encriptado
Public gblnPausar As Boolean    'Pausar & Reactivar
Public giIdioma As Integer      'Para saber el Idioma

Public MyDate As Date         'Fecha interna


