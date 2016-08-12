# DEFINE DEBUGMODE .F.
*On Error Do ErrorHandler With Error( ), ;
							Message( ), ;
							Message(1), ;
							Program( ), ;
							Lineno(1)

ON SHUTDOWN DO ExitProg

RELEASE ALL EXTENDED
CLEAR ALL
SET TALK OFF
SET HOURS TO 24
SET DATE TO GERMAN
SET CENTURY ON 
SET CONSOLE OFF
SET RESOURCE OFF 
SET safety OFF 
SET REPROCESS TO 3 SECONDS 
SET DELETED ON 
SET ESCAPE OFF  

WITH _SCREEN
 .Width      = 1024
 .Height     = 768-100
 .BackColor  = RGB(192,192,192)
 .AutoCenter = .f.
ENDWITH 

DECLARE INTEGER GetPrivateProfileString IN Win32API  AS GetPrivStr ;
	STRING cSection, STRING cKey, STRING cDefault, STRING @cBuffer, ;
	INTEGER nBufferSize, STRING cINIFile
DECLARE INTEGER WritePrivateProfileString IN Win32API AS WritePrivStr ;
	STRING cSection, STRING cKey, STRING cValue, STRING cINIFile
DECLARE INTEGER GetSysColor IN User32.DLL INTEGER
DECLARE INTEGER ShellExecute IN shell32.dll ;
	INTEGER hndWin, STRING cAction, STRING cFileName, ;
	STRING cParams, STRING cDir, INTEGER nShowWin

PUBLIC fso AS SCRIPTING.FileSystemObject, wshell AS Shell.Application

fso    = CREATEOBJECT('Scripting.FileSystemObject')
WShell = CREATEOBJECT('Shell.Application')
WSHShell = CREATEOBJECT('Wscript.Shell')

*SET ENGINEBEHAVIOR 70

SET PROCEDURE TO Utils

*application.visible = .f.
_SCREEN.Picture = 'attppl.jpg'
_screen.Visible=.t.

lcPathSystem = sys(5) + sys(2003)
Set DEFAULT TO (lcPathSystem)
lcPathMain = sys(5) + sys(2003)
lcPathSystem = lcPathMain+;
	','+(lcPathMain+'\BITMAPS')+;
	','+(lcPathMain+'\FORMS')+;
	','+(lcPathMain+'\GRAPHICS')+;
	','+(lcPathMain+'\INCLUDE')+;
	','+(lcPathMain+'\LIBS')+;
	','+(lcPathMain+'\MENU')+;
	','+(lcPathMain+'\PROGRAM')

SET PATH TO (lcPathSystem)

PUBLIC paisoms, parc, pbase, pbin, pcommon, pout, ptempl, ptrash, pdouble, pmee, DaemonDir, SpamDir, ;
 tyear, tmonth, tdat1, tdat2, pgdat1, pgdat2, curlpu, qcod, qmail, qobjid, UserERZ, qname, oMenu, gcPeriod, pcPeriod, gcUser, gcFormat,;
 usrfam, usrim, usrot, usrfio

PUBLIC ARRAY mes_text[12], mes_main[12]
mes_text[1]="января"
mes_text[2]="февраля"
mes_text[3]="марта"
mes_text[4]="апреля"
mes_text[5]="мая"
mes_text[6]="июня"
mes_text[7]="июля"
mes_text[8]="августа"
mes_text[9]="сентября"
mes_text[10]="октября"
mes_text[11]="ноября"
mes_text[12]="декабря"

mes_main[1]="Январь"
mes_main[2]="Февраль"
mes_main[3]="Март"
mes_main[4]="Апрель"
mes_main[5]="Май"
mes_main[6]="Июнь"
mes_main[7]="Июль"
mes_main[8]="Август"
mes_main[9]="Сентябрь"
mes_main[10]="Октябрь"
mes_main[11]="Ноябрь"
mes_main[12]="Декабрь"

IF CfgBase() = -1
 =ExitProg()
ENDIF 

IF !fso.FolderExists(pcommon)
 MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТСТВУЕТ ИЛИ НЕДОСТУПНА'+CHR(13)+CHR(10)+'ДИРЕКТОРИЯ '+pcommon,0+16,'')
 =ExitProg()
ENDIF  

m.tdat1 = CTOD('01.'+PADL(tMonth,2,'0')+'.'+PADL(tYear,4,'0'))
m.tdat2 = GOMONTH(CTOD('01.'+PADL(tMonth,2,'0')+'.'+PADL(tYear,4,'0')),1)-1
m.pgdat1 = m.tdat1
m.pgdat2 = m.tdat2

m.gcPeriod = STR(tYear,4)+PADL(tMonth,2,'0')
m.pcPeriod = STR(IIF(tMonth!=1, tYear, tYear-1),4)+PADL(IIF(tMonth!=1,tMonth-1,12),2,'0')

DO CASE 
 CASE m.qcod == 'P2'
  m.qname = 'ОАО "МЕДИЦИНСКАЯ СТРАХОВАЯ КОМПАНИЯ "УРАЛСИБ"'
  m.qmail = 'skpomed.msk.oms'
  m.qobjid = 3386
 CASE m.qcod == 'P3'
  m.qname = 'ООО МСО "ПАНАЦЕЯ" МОСКОВСКИЙ ФИЛИАЛ'
  m.qmail = 'panacea.msk.oms'
  m.qobjid = 5387
 CASE m.qcod == 'I3'
  m.qname = 'ООО СК "ИНГОССТРАХ-М"'
  m.qmail = 'ingos.msk.oms'
  m.qobjid = 5398
 CASE m.qcod == 'I1'
  m.qname = 'ООО МСК "ИКАР"'
  m.qmail = 'ikar.msk.oms'
  m.qobjid = 110
 CASE m.qcod == 'R4'
  m.qname = 'ООО "СТРАХОВАЯ МЕДИЦИНСКАЯ КОМПАНИЯ РЕСО-МЕД" МОСКОВСКИЙ ФИЛИАЛ'
  m.qmail = 'reso.msk.oms'
  m.qobjid = 3415
 CASE m.qcod == 'S2'
  m.qname = 'АО ВТБ Медицинское страхование'
  m.qmail = 'sovita.msk.oms'
  m.qobjid = 33
 CASE m.qcod == 'S7'
  m.qname = 'АО СК "СОГАЗ-Мед"'
  m.qmail = 'sogaz.msk.oms'
  m.qobjid = 5400
 CASE m.qcod == 'S6'
  m.qname = 'ЗАО СК "СОГЛАСИЕ-М"'
  m.qmail = 'soglasie.msk.oms'
  m.qobjid = 5404
 CASE m.qcod == 'R8'
  m.qname = 'РГС'
  m.qmail = 'rgs.msk.oms'
  m.qobjid = 5405
 OTHERWISE 
  m.qname = 'ОАО "МЕДИЦИНСКАЯ СТРАХОВАЯ КОМПАНИЯ "УРАЛСИБ"'
  m.qmail = 'skpomed.msk.oms'
  m.qobjid = 3386
ENDCASE 

public goApp
goApp = NEWOBJECT('_goapp','main')
ADDPROPERTY(goApp, "mcod", "")
ADDPROPERTY(goApp, "lpuid", 0)

goApp.Begin_process()

=chkbase()

=OpenFile(pCommon+'\Users', 'Users', 'shar', 'name')
SELECT Users
IF !SEEK(m.gcUser, 'Users')
 USE 
 MESSAGEBOX('ИМЯ '+ALLTRIM(m.gcUser)+' ОТСУТСТВУЕТ В СПРАВОЧНИКЕ!', 0+16, '')
 =ExitProg()
ELSE 
 IF !RLOCK()
  USE 
  MESSAGEBOX('ПОЛЬЗОВАТЕЛЬ '+ALLTRIM(m.gcUser)+' УЖЕ РАБОТАЕТ В СИСТЕМЕ!', 0+16, '')
  =ExitProg()
 ELSE 
  m.usrfam = ALLTRIM(Fam)
  m.usrim  = ALLTRIM(im)
  m.usrot  = ALLTRIM(ot)
  m.usrfio = ALLTRIM(fio)
 ENDIF 
ENDIF 


=OpenFile(pCommon+'\smo', 'smo', 'shar')
SELECT smo
COUNT FOR v != .f. TO kol_q
PUBLIC smo(kol_q, 2)
COPY FOR v != .f. FIELDS code, name TO ARRAY smo
USE 

numlib = adir(alib,lcPathMain+'\LIBS\*.vcx')
for i = 1 to numlib
	lcSetLibrary = lcPathMain+'\LIBS\' + alib(i,1)
	set classlib to (lcSetLibrary) additive
endfor

SET LIBRARY TO &lcPathMain\vfpzip
SET LIBRARY TO &lcPathMain\vfpexmapi ADDITIVE

SET SYSMENU TO
SET SYSMENU ON
SET STATUS BAR ON 
WITH _SCREEN
 .Icon = 'cross.ico'
 .Caption = m.qname+', ПОЛЬЗОВАТЕЛЬ: '+ALLTRIM(m.gcUser)+', ПЕРИОД: '+NameOfMonth(tMonth)+' '+STR(tYear,4)+' ГОДА'+;
  ' (с '+DTOC(tdat1)+' по '+DTOC(tdat2)+')'
ENDWITH 

*DO DelSpareFiles

DO m_menu
 
READ EVENTS

=ExitProg()


FUNCTION ExitProg
 IF USED('Users')
  USE IN Users
 ENDIF 
 RELEASE m.oError && ???
 #IF DEBUGMODE
*  _SCREEN.Caption = oApp.cOldWindCaption
  SET SYSMENU TO DEFAULT
  _SCREEN.TitleBar = 1
  _SCREEN.WindowState = 2
  _SCREEN.LockScreen = .F.
  _SCREEN.Picture = ''
  _SCREEN.BackColor = RGB(255,255,255)
*  oApp.ShowToolBars()
  SET SYSMENU ON
 #ELSE
 _SCREEN.Caption = ""
 #ENDIF
* oApp.CloseAllTable()
* RELEASE m.oApp
 RELEASE m.goApp
 #IF !DEBUGMODE
  ON SHUTDOWN
  QUIT
 #ELSE
  ON SHUTDOWN
  _SCREEN.Icon =""
  _SCREEN.FirstStart = .T.
  *SET HELP TO
  CLEAR ALL
  CLOSE ALL
  CLEAR PROGRAM
  SET SYSMENU NOSAVE
  SET SYSMENU TO DEFAULT
  SET SYSMENU ON
 #ENDIF
RETURN 
