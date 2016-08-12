PROCEDURE LoadErrFomsN

lcUser = 'OMS'

IF MESSAGEBOX(CHR(13)+CHR(10)+'ПОИСКАТЬ ОШИБКИ МГФОМС?'+CHR(13)+CHR(10),4+32,'')==7
 RETURN 
ENDIF 
IF !IsAisDir() && Проверка наличия директорий, OMS, INPUT, OUTPUT
 RETURN 
ENDIF 

oMailDir        = fso.GetFolder(pAisOms+'\&lcUser\input')
MailDirName     = oMailDir.Path
oFilesInMailDir = oMailDir.Files
nFilesInMailDir = oFilesInMailDir.Count

IF nFilesInMailDir<=0
 MESSAGEBOX(CHR(13)+CHR(10)+'В ДИРЕКТОРИИ АИСОМС НЕТ НЕПРИНЯТЫХ ФАЙЛОВ!'+CHR(13)+CHR(10),0+64,'')
 RETURN 
ENDIF 

WAIT "ПРОСМОТР ПОЧТЫ..." WINDOW NOWAIT 

OldEscStatus = SET("Escape")
SET ESCAPE OFF 
CLEAR TYPEAHEAD 

FOR EACH oFileInMailDir IN oFilesInMailDir

 IF LOWER(oFileInMailDir.Name) != 'b'
  LOOP 
 ENDIF 

 m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)

 m.BFullName = oFileInMailDir.Path
 m.bname     = oFileInMailDir.Name
 m.recieved  = oFileInMailDir.DateLastModified
 m.lpuid     = 0
 m.processed = DATETIME()
 
 m.cfrom      = ''
 m.cdate      = ''
 m.cmessage   = ''
 m.resmesid   = ''
 m.csubject   = ''
 m.csubject1  = ''
 m.csubject2  = ''
 m.attachment = ''
 m.bodypart   = ''

 m.attaches   = 0 && Сколько присоединенных файлов в одной ИП
 DIMENSION dattaches(10,2)
 dattaches = ''

 m.bparts   = 0 && Сколько присоединенных файлов в одной ИП
 DIMENSION dbparts(10,2)
 dbparts = ''

 CFG = FOPEN(m.BFullName)
 =ReadCFGFile()
 =FCLOSE (CFG)

 IF UPPER(m.cfrom)!='OMS@MGF.MSK.OMS'
  LOOP 
 ENDIF 
 IF UPPER(RIGHT((ALLTRIM(m.csubject)),2))!='GP'
  LOOP 
 ENDIF 
 IF m.attaches!=1
  LOOP 
 ENDIF 
 IF LEFT(UPPER(dattaches(m.attaches,2)),3)!='K'+m.qcod
  LOOP 
 ENDIF 

 m.sent = dt2date(m.cdate)
 WAIT m.cfrom WINDOW NOWAIT 
   
 m.llIsSubject = .F.

 && Присоединено ли что-нибудь к файлу? Если нет, то - в спам!
 IF m.attaches == 0 AND m.bparts == 0 && Если к файлу ничего не присоединено!
  LOOP 
 ENDIF 
 && Присоединено ли что-нибудь к файлу? Если нет, то - в спам!

 && Проверка комплектности посылки
 IsComplect = .T.
 IF m.attaches>0
  FOR natt = 1 TO m.attaches
   IF !fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    IsComplect = .F.
    LOOP 
   ENDIF 
  ENDFOR 
 ENDIF 
 IF IsComplect = .F.
  LOOP 
 ENDIF  

 IsComplect = .T.
 IF m.bparts>0
  FOR natt = 1 TO m.bparts
   IF !fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    IsComplect = .F.
    LOOP 
   ENDIF 
  ENDFOR 
 ENDIF 
 IF IsComplect = .F.
  LOOP 
 ENDIF  
 && Проверка комплектности посылки
 
 && Если это нормальная и новая посылка!
 InDirPeriod = pBase + '\' + m.gcPeriod
 IF !fso.FolderExists(InDirPeriod)
  LOOP 
 ENDIF
 InDir = InDirPeriod

* fso.CopyFile(m.BFullName, InDir + '\' + m.bname)
 fso.CopyFile(m.BFullName, pArc + '\IPS\INPUT\' + m.bname)
 fso.DeleteFile(m.BFullName)
 IF m.attaches>0
  FOR nattach = 1 TO m.attaches
   m.ddname   = ALLTRIM(dattaches(nattach, 1))
   m.aattname = ALLTRIM(dattaches(nattach, 2))
   IF !EMPTY(m.ddname)

    ffile = fso.GetFile(pAisOms+'\'+lcUser+'\input\'+m.ddname)
    IF ffile.size >= 2
     fhandl = ffile.OpenAsTextStream
     lcHead = fhandl.Read(2)
     fhandl.Close
    ELSE 
     lcHead = ''
    ENDIF 

    IF lcHead == 'PK' && Это zip-файл!
     ZipName = pAisOms+'\'+lcUser+'\input\'+m.ddname
     UnzipOpen(ZipName)
     erItem   = 'ER' + m.qcod + PADL(m.tmonth,2,'0')+SUBSTR(STR(m.tyear,4),3,2)+'.dbf'
     IF UnzipGotoFileByName(erItem)
      llIsOneZip = .t.
      UnzipClose()
     ENDIF 
     UnzipClose()
    ELSE 
     LOOP 
    ENDIF 

    IF llIsOneZip == .F.
     LOOP 
    ENDIF 

    fso.CopyFile(MailDirName + '\' + m.ddname, InDir+'\'+LOWER(m.aattname), .t.)
*    fso.CopyFile(MailDirName + '\' + m.ddname, InDir+'\'+m.ddname, .t.)
    fso.CopyFile(MailDirName + '\' + m.ddname, parc+'\IPS\INPUT\'+m.ddname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.ddname)
    
    IF fso.FileExists(InDir+'\'+LOWER(m.aattname))
     MESSAGEBOX(InDir+'\'+LOWER(m.aattname))
     ZipName = InDir+'\'+LOWER(m.aattname)
     UnzipOpen(ZipName)
     UnzipGotoFileByName(erItem)
     UnzipFile(InDir+'\')
     UnzipClose()
    ENDIF 

   ENDIF 
  ENDFOR 
 ENDIF 
 IF m.bparts > 0
  FOR npart = 1 TO m.bparts
   m.bpname   = ALLTRIM(dbparts(npart, 1))
   IF !EMPTY(m.bpname)
    fso.CopyFile(MailDirName + '\' + m.bpname, InDir+'\'+m.bpname, .t.)
    fso.CopyFile(MailDirName + '\' + m.bpname, InDir+'\'+m.bpname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.bpname)
   ENDIF 
  ENDFOR 
 ENDIF 

 IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 

NEXT && Цикл по файлам

SET ESCAPE &OldEscStatus

WAIT CLEAR 

MESSAGEBOX(CHR(13)+CHR(10)+'ОБРАБОТКА ЗАКОНЧЕНА!'+CHR(13)+CHR(10),0+64,'')

RETURN 

FUNCTION ReadCFGFile
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  DO CASE
   CASE UPPER(READCFG) = 'FROM'
    m.cfrom = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'DATE'
    m.cdate = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'MESSAGE'
    m.cmessage = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'RESENT-MESSAGE-ID'
    m.resmesid = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'SUBJECT'
    m.csubject = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    m.csubject1 = LEFT(m.csubject, RAT('#',m.csubject,2))   && Делим subject для последующей вставки кода результата
    m.csubject2 = SUBSTR(m.csubject, RAT('#',m.csubject,1)) && Делим subject для последующей вставки кода результата
   CASE UPPER(READCFG) = 'ATTACHMENT'
    m.attaches   = m.attaches + 1
    m.attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dattaches(m.attaches,1) = ALLTRIM(SUBSTR(m.attachment, 1, AT(" ",m.attachment)-1)) && Название d-файла
    dattaches(m.attaches,2) = ALLTRIM(SUBSTR(m.attachment, AT(" ",m.attachment)+1))    && Фактическое название файла
   CASE UPPER(READCFG) = 'BODYPART'
    m.bparts   = m.bparts + 1
    m.bodypart = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dbparts(m.bparts,1) = ALLTRIM(SUBSTR(m.bodypart, 1, AT(" ",m.bodypart)-1))
  ENDCASE
 ENDDO
RETURN 

FUNCTION IsAisDir()
 IF !fso.FolderExists(pAisOms)
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms, 0+16, '')
  RETURN .F.
 ENDIF 
 IF !fso.FolderExists(pAisOms+'\&lcUser')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser', 0+16, '')
  RETURN .F.
 ENDIF 
 IF !fso.FolderExists(pAisOms+'\&lcUser\INPUT')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser\INPUT', 0+16, '')
  RETURN .F.
 ENDIF 
 IF !fso.FolderExists(pAisOms+'\&lcUser\OUTPUT')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser\OUTPUT', 0+16, '')
  RETURN .F. 
 ENDIF
RETURN .T. 
