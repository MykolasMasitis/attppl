PROCEDURE CheckZip
PARAMETERS lcUser

IF !IsAisDir() && Проверка наличия директорий, OMS, INPUT, OUTPUT
 RETURN 
ENDIF 

oMailDir        = fso.GetFolder(pAisOms+'\&lcUser')
MailDirName     = oMailDir.Path
oFilesInMailDir = oMailDir.Files
nFilesInMailDir = oFilesInMailDir.Count

MESSAGEBOX('ОБНАРУЖЕНО '+ALLTRIM(STR(nFilesInMailDir))+' ФАЙЛОВ!', 0+64, lcUser)

IF nFilesInMailDir<=0
 RETURN 
ENDIF 

IF OpenTemplates() != 0
 =CloseTemplates() 
 RETURN 
ENDIF 

WAIT "ПРОСМОТР ПОЧТЫ..." WINDOW NOWAIT 
SELECT AisOms
prvorder = ORDER('aisoms')
SET ORDER TO 

m.un_id = SYS(3)

OldEscStatus = SET("Escape")
SET ESCAPE OFF 
CLEAR TYPEAHEAD 

FOR EACH oFileInMailDir IN oFilesInMailDir

 SCATTER MEMVAR BLANK
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

 DO CASE 
 CASE LEFT(UPPER(oFileInMailDir.Name),2) = 'PR'
 m.lpuid = INT(VAL(SUBSTR(oFileInMailDir.Name,3,4)))
 IF m.lpuid==0
  LOOP 
 ENDIF 
 m.adresat = IIF(SEEK(m.lpuid, "sprabo", "lpu_id"), ALLTRIM(sprabo.abn_name), '')
 m.cfrom = IIF(SEEK(m.lpuid, "sprabo", "lpu_id"), 'oms@'+ m.adresat, '')

 m.sent = DATETIME()
 WAIT m.cfrom WINDOW NOWAIT 
   
 m.llIsSubject = .F.
 
 m.headid     = 0
 m.headname   =  ''
 m.branchid   = 0
 m.branchname =  ''
 m.ishead     = .f.
 
 m.headid     = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.lpu_id, 0)
 m.headname   = IIF(SEEK(m.lpuid, "sprlpu"), ALLTRIM(sprlpu.fullname), '')
 m.headmcod   = IIF(SEEK(m.headid, "sprlpu", 'fil_id'), sprlpu.mcod, "")
 m.branchid   = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.fil_id, 0)
 m.branchname = IIF(SEEK(m.lpuid, "sprlpu"), ALLTRIM(sprlpu.name), '')
 
 m.ishead = IIF(m.headid==m.branchid, .t., .f.)
 
 IF !m.ishead
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ПОЛУЧЕНО ИП ОТ ЛПУ'+CHR(13)+CHR(10)+;
   m.branchname+' ,'+CHR(13)+CHR(10)+;
   'РЕОРГАНИЗОВАННОГО В ФИЛИАЛ АПО'+CHR(13)+CHR(10)+;
   m.headname+CHR(13)+CHR(10)+;
   'ПРИНИМАТЬ К ОБРАБОТКЕ ИНФОРМАЦИЮ?'+CHR(13)+CHR(10),4+16,'') == 7
   LOOP  
  ELSE 
   IF !fso.FolderExists(pbase+'\'+gcperiod+'\'+m.headmcod)
    MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ ГОЛОВНОЙ ОРГАНИЗАЦИИ'+CHR(13)+CHR(10)+;
     '('+pbase+'\'+gcperiod+'\'+m.headmcod+')'+CHR(13)+CHR(10),0+64,'')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+gcperiod+'\'+m.headmcod+'\people.dbf')
    MESSAGEBOX(CHR(13)+CHR(10)+'ПРИНЯТЬ ИП ОТ ФИЛИАЛА МОЖНО ТОЛЬКО'+CHR(13)+CHR(10)+;
     'ПОСЛЕ ПРИЕМА ИП ОТ ГОЛОВНОЙ ОРГАНИЗАЦИИ!'+CHR(13)+CHR(10),0+64,m.headmcod)
    LOOP 
   ENDIF 
  ENDIF 
 ENDIF 
 
 IF m.ishead
  m.mcod    = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.mcod, "")
 ELSE 
  m.mcod = m.headmcod
 ENDIF 

 IF EMPTY(m.mcod)
  LOOP 
 ENDIF 

 m.cokr    = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.cokr, "")
 m.moname  = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.name, "")
 m.usr     = IIF(SEEK(m.lpuid, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")

 IF EMPTY(m.usr) AND m.gcUser!='OMS'
  LOOP 
 ENDIF 
 
 IF m.usr != m.gcUser AND m.gcUser!='OMS'
  LOOP 
 ENDIF 

 m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)

 m.previous_id = m.un_id
 m.un_id     = SYS(3)
 DO WHILE m.un_id = m.previous_id
  m.un_id     = SYS(3)
 ENDDO 
 m.previous_id = m.un_id

 m.tansfile  = 'tok_' + m.mcod
 m.bansfile  = 'bok_' + m.mcod
 iii = 1
 DO WHILE fso.FileExists(pAisOms+'\OMS\OUTPUT\'+m.bansfile)
  m.tansfile  = 'tok_' + m.mcod + '_' + PADL(iii,2,'0')
  m.bansfile  = 'bok_' + m.mcod + '_' + PADL(iii,2,'0')
  iii = iii + 1
 ENDDO 

 m.messageid = ALLTRIM(m.un_id+'.'+m.gcUser+'@'+m.qmail)

 llIsOneZip = .F.

 ffile = fso.GetFile(m.BFullName)
 IF ffile.size >= 2
  fhandl = ffile.OpenAsTextStream
  lcHead = fhandl.Read(2)
  fhandl.Close
 ELSE 
  lcHead = ''
 ENDIF 

 IF lcHead == 'PK' && Это zip-файл!
  ZipName = m.BFullName
  IsSuccess=UnzipOpen(ZipName)
  IF IsSuccess
   llIsOneZip = .t.
   UnzipClose()
  ENDIF 
  UnzipClose()
 ENDIF 
 
 IF llIsOneZip == .F.
  LOOP 
 ENDIF 

 && Проверяем комплектность посылки - наличие 1 файла!
 UnzipOpen(ZipName)
 rItem    = 'PR'+STR(m.lpuid,4)+m.qcod + '.'+m.mmy
 IF !IsIPComplete()
  LOOP 
 ENDIF 
 UnzipClose()
 && Проверяем комплектность посылки - наличие 1 файла!

 && Если посылка повторная
 IsThisIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod') AND !EMPTY(AisOms.Sent), .t., .f.)
 IF IsThisIPDouble AND m.IsHead
  LOOP 
 ENDIF 
 && Если посылка повторная

 && Если это нормальная и новая посылка!
 InDirPeriod = pBase + '\' + m.gcPeriod
 IF !fso.FolderExists(InDirPeriod)
  fso.CreateFolder(InDirPeriod)
 ENDIF 
 InDir = pBase + '\' + m.gcPeriod + '\' + m.mcod
 IF !fso.FolderExists(InDir)
  fso.CreateFolder(InDir)
 ENDIF 

 fso.CopyFile(m.bfullname, InDir+'\'+m.bname, .t.)
 fso.CopyFile(m.bfullname, parc+'\IPS\INPUT\'+m.bname, .t.)
 fso.DeleteFile(m.bfullname)

 ZipName = InDir + '\' + m.bname
 ZipDir  = InDir + '\'

 UnzipOpen(ZipName)
 UnzipGotoFileByName(rItem)
 UnzipFile(ZipDir)
 UnzipClose()

 m.lcCurDir = pBase + '\' + m.gcPeriod + '\' + m.mcod+'\'
 SET DEFAULT TO (lcCurDir)

 =OpenFile("&rItem",  "rfile",  "SHARED")
 
 IF !CheckFilesStucture()
  LOOP 
 ENDIF 
 
 =CloseItems()

 lcDir  = m.pBase + '\' + m.gcPeriod + '\' + m.mcod
 People = lcDir + '\people'
 Error  = lcDir + '\e' + m.mcod

* IF m.IsHead
*  =CreateFilesStructure()
* ENDIF 

* =OpenLocalFiles()
* USE (People) IN 0 ALIAS People SHARED
 =OpenFile(People, 'People', 'Shar')
 
 m.paztot = 0 
 m.pazin  = 0
* IF m.IsHead
*  =MakePeople()
* ELSE 
  =OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+ritem, 'ritem', 'shar')
*  =OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'n_pol')
  SELECT ritem
  m.paztot = paztot
  m.pazin  = pazin
  SCAN 
   SCATTER MEMVAR
   m.recid_lpu = m.recid
   m.ddr = CTOD(SUBSTR(m.dr,7,2)+'.'+SUBSTR(m.dr,5,2)+'.'+LEFT(m.dr,4))
   RELEASE m.recid, m.dr, m.q, m.lpu_id
   m.dr = m.ddr
   RELEASE m.ddr
   IF !SEEK(m.n_pol, 'people', 'n_pol')
    m.paztot = m.paztot + 1
    m.pazin  = m.pazin + IIF(EMPTY(date_out),1,0)
    INSERT INTO People FROM MEMVAR
   ELSE 
    UPDATE People SET recid_lpu=m.recid_lpu, date_in=m.date_in, date_out=m.date_out, spos=m.spos, s_pol=m.s_pol, ;
     n_pol=m.n_pol, tip_d=m.tip_d, fam=m.fam, im=m.im, ot=m.ot, dr=m.dr, w=m.w WHERE n_pol=m.n_pol
   ENDIF 
  ENDSCAN 
  USE IN ritem
  USE IN people
* ENDIF 

 MailView.recs   = MailView.recs   + 1
 MailView.paztot = MailView.paztot + m.paztot
 MailView.pazin  = MailView.pazin  + m.pazin

 =SEEK(m.mcod, 'aisoms', 'mcod')
 MailView.refresh

 m.opaztot = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.paztot, 0)
 m.opazin  = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.pazin, 0)

* IF m.ishead
*  UPDATE aisoms SET bname=m.bname, dname=m.dname, sent=m.sent, recieved=m.recieved, ;
   processed=m.processed, cfrom=m.cfrom, cmessage=m.cmessage, erzsv4_id='', sv4st=0, ;
   erzlpu_id='', lpust=0, paztot=m.paztot, pazin=m.pazin, childs=0, adults=0;
   WHERE mcod=m.mcod
  UPDATE aisoms SET bname=m.bname, dname=m.dname, sent=m.sent, recieved=m.recieved, ;
   processed=m.processed, cfrom=m.cfrom, cmessage=m.cmessage, erzsv4_id='', sv4st=0, ;
   erzlpu_id='', lpust=0, paztot = m.opaztot+m.paztot, pazin=m.opazin+m.pazin, childs=0, adults=0;
   WHERE mcod=m.mcod
* ELSE 
*  m.opaztot = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.paztot, 0)
*  m.opazin  = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.pazin, 0)

*  UPDATE aisoms SET erzsv4_id='', sv4st=0, erzlpu_id='', lpust=0, childs=0, adults=0,;
   paztot=m.opaztot+m.paztot, pazin=m.opazin+m.pazin;
   WHERE mcod=m.mcod
* ENDIF 

 SELECT AisOms
 
 WAIT CLEAR 
 MailView.Refresh 

 OTHERWISE 

 ENDCASE 

 IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 

NEXT && Цикл по файлам

SET ESCAPE &OldEscStatus

=CloseTemplates() 

SET ORDER TO (prvorder)
MailView.Refresh
MailView.LockScreen=.f.

WAIT CLEAR 

nFilesInMailDir = oFilesInMailDir.Count

MESSAGEBOX('ОСТАЛОСЬ '+ALLTRIM(STR(nFilesInMailDir))+' НЕОБРАБОТАННЫХ ИП!', 0+64, lcUser)

RETURN 

FUNCTION CopyToTrash(lcPath, nTip)
 fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pAisOms+'\&lcUser\OUTPUT\'+m.bansfile)
 fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pArc+'\IPS\OUTPUT\'+m.bansfile)
 fso.DeleteFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile)
 TrashDir = pTrash + '\' + m.mcod
 IF !fso.FolderExists(TrashDir)
  fso.CreateFolder(TrashDir)
 ENDIF 
 fso.CopyFile(lcPath + '\' + m.bname, TrashDir+'\'+m.bname)
 fso.DeleteFile(lcPath + '\' + m.bname)

 FOR nattach = 1 TO m.attaches
  IF !EMPTY(ALLTRIM(dattaches(nattach, 1)))
   fso.CopyFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)), TrashDir + '\' + ALLTRIM(dattaches(nattach, 2)))
   fso.DeleteFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)))
  ENDIF 
 ENDFOR 

 IF NOT SEEK(m.cmessage, "taisoms", "cmessage")
  INSERT INTO taisoms FROM MEMVAR 
 ENDIF
RETURN 

FUNCTION ClDir
 IF fso.FileExists(rItem)
  DELETE FILE &ritem
 ENDIF 
RETURN 

FUNCTION OpenTemplates
 tn_result = 0
 tn_result = tn_result + OpenFile("&ptempl\prxxxxqq.mmy", "pr_et", "SHARED")
RETURN tn_result

FUNCTION CloseTemplates
 IF USED('pr_et')
  USE IN pr_et
 ENDIF 
RETURN 

FUNCTION CloseItems
 IF USED('rfile')
  USE IN rfile
 ENDIF 
RETURN 

FUNCTION CompFields(NameOfFile)
 FOR nFld = 1 TO fld_1
  IF (tabl_1(nFld,1) == tabl_2(nFld,1)) AND ;
     (tabl_1(nFld,2) == tabl_2(nFld,2)) AND ;
     (tabl_1(nFld,3) == tabl_2(nFld,3))
  ELSE 
   =CloseItems()
*   =ClDir()
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Wrong structure of ] + NameOfFile
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   RETURN 0 
  ENDIF 
 ENDFOR 
RETURN 1

FUNCTION DiffFields(NameOfFile)
 =CloseItems()
* =ClDir()
 m.csubject = m.csubject1 + '08' +m.csubject2
 m.cerrmessage = [Wrong number of fields in ] + NameOfFile
 IF m.llIsSubject = .F.
  m.llIsSubject = .T.
  poi.WriteLine('Subject: '+m.csubject)
 ENDIF 
 poi.WriteLine('BodyPart: ' + m.cerrmessage)
 poi.Close
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

FUNCTION WriteInBFile(BFullName, TextToWrite)
 CFG = FOPEN(BFullName,12)
 IsMyCommentExists = .F.
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  IF UPPER(READCFG) = 'MYCOMMENT'
   IsMyCommentExists = .T.
   LOOP 
  ENDIF 
 ENDDO
 IF !IsMyCommentExists
  nFileSize = FSEEK(CFG,0,2)
  =FWRITE(CFG, TextToWrite)
 ENDIF 
 = FCLOSE (CFG)
RETURN 

FUNCTION MakePeople
 tFile = lcDir+'\'+rItem
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&rItem IN 0 ALIAS lcRFile  EXCLUSIVE 
 SELECT lcRFile
 SCAN 
  SCATTER MEMVAR
  m.recid_lpu = m.recid
  m.ddr = CTOD(SUBSTR(m.dr,7,2)+'.'+SUBSTR(m.dr,5,2)+'.'+LEFT(m.dr,4))
  RELEASE m.recid, m.dr, m.q, m.lpu_id
  m.dr = m.ddr
  RELEASE m.ddr
  m.paztot = m.paztot + 1
  m.pazin  = m.pazin + IIF(EMPTY(date_out),1,0)
  INSERT INTO People FROM MEMVAR
 ENDSCAN 
 USE 
 USE IN people 
 fso.DeleteFile(lcDir+'\'+rItem)
RETURN 

FUNCTION CreateFilesStructure

 CREATE TABLE (People) ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1, recid_lpu c(6), lpu_id n(6), date_in d, date_in2 d, date_out d, ;
   spos c(1), spos2 c(1), s_pol c(6), s_pol2 c(6), n_pol c(16), n_pol2 c(16), tip_d c(1), tip_d2 c(1), ;
   q c(2), fam c(25), im c(20), ot c(20), dr d, w n(1), c_err c(2), ans_r c(3), lpuid_et n(6), datein_et d,;
   spos_et c(1), reserv c(20))
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON n_pol TAG n_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX ON dr TAG dr
 INDEX ON c_err TAG c_err
 USE 

RETURN 

FUNCTION OpenLocalFiles
 USE (People) IN 0 ALIAS People SHARED
RETURN 

FUNCTION CheckFilesStucture

 fld_1 = AFIELDS(tabl_1, 'rfile') && проверка r-файла
 fld_2 = AFIELDS(tabl_2, 'pr_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(rItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(rItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

RETURN .T. 

FUNCTION IsIPComplete
 DO CASE 
  CASE !UnzipGotoFileByName(rItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Отсутствует ] + rItem + [ файл]
   UnzipClose()
   m.cmnt = m.cerrmessage
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   =CopyToTrash(m.MailDirName,1)
   RETURN .F. 
 ENDCASE
RETURN .T.

