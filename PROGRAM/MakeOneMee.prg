FUNCTION MakeOneMEE(lcPath, IsVisible, IsQuit, TipAcc, TipOfExp) && Плановая МЭЭ
 
 m.TipOfExp = TipOfExp

 DO CASE 
  CASE m.TipOfExp = '2'
   m.ctipofexp = 'плановой'
  CASE m.TipOfExp = '3'
   m.ctipofexp = 'целевой'
  OTHERWISE 
   m.ctipofexp = ''
 ENDCASE 

 oal = ALIAS()
 
 m.lWasUsedSprLpu = IIF(USED('sprlpu'), .T., .F.)
 m.lWasUsedTarif  = IIF(USED('tarif'), .T., .F.)
 m.lWasUsedMkb    = IIF(USED('mkb10'), .T., .F.)
 m.lWasUsedSooKod = IIF(USED('sookod'), .T., .F.)
 m.lWasUsedTalon  = IIF(USED('talon'), .T., .F.)
 m.lWasUsedPeople = IIF(USED('people'), .T., .F.)
 m.lWasUsedDoctor = IIF(USED('doctor'), .T., .F.)
 m.lWasUsedSSActs = IIF(USED('ssacts'), .T., .F.)
 m.lWasUsedMError = IIF(USED('merror'), .T., .F.)

 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF !OpenFiles(lcpath)
  RETURN 
 ENDIF 
 
 DotName = 'Акт_МЭЭ_СС.dot'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ ШАБЛОН ОТЧЕТА'+CHR(13)+CHR(10)+;
   'Акт_МЭЭ_СС.dot',0+32,'')
  RETURN 
 ENDIF 
 
 m.mcod    = mcod 
 m.lpuid   = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.name)+', '+m.mcod, '')
 m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)

 m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)

 m.sn_pol = sn_pol
 m.ppolis = STRTRAN(ALLTRIM(m.sn_pol),' ','') && Для названия Акта
 
 oldalias = ALIAS()
 orecno = RECNO()
* SELECT merror
 SELECT talon 
 
 COUNT FOR sn_pol = m.sn_pol AND SEEK(recid, 'merror', 'recid') AND merror.et = m.TipOfExp TO m.nIsExps

 IF m.nIsExps == 0
  MESSAGEBOX(CHR(13)+CHR(10)+'ПО ВЫБРАННОМУ СЧЕТУ МЭЭ НЕ ПРОВОДИЛОСЬ!'+CHR(13)+CHR(10),0+64,'')
  CloseFiles()
  SELECT (oldalias)
  GO (orecno)
  RETURN
 ELSE 
 ENDIF 

 CREATE CURSOR ttalon (sn_pol c(25),c_i c(25),ds c(6),tip c(1),d_u d,pcod c(10),otd c(4),cod n(6),k_u n(3),;
  d_type c(1),s_all n(11,2),e_cod n(6),e_ku n(3),e_tip c(1),err_mee c(3),e_period c(6),et c(1),e_sall n(11,2),;
  koeff n(4,2), ee c(1), straf n(4,2))

 SELECT talon 

 m.TipOfMee = 0
 
 m.s_lech = 0
 m.dat1 = {31.12.2099}
 m.dat2 = {01.01.2000}
 m.dslast = 0
 SCAN FOR sn_pol == m.sn_pol

  IF SEEK(recid, 'serror')
   LOOP 
  ENDIF 
  IF !SEEK(PADL(recid,6,'0')+m.TipOfExp, 'merror', 'id_et')
   LOOP 
  ENDIF 

  SCATTER FIELDS EXCEPT recid,mcod,period,sn_pol,q,novor,ds_s,ds_p,profil,rslt,prvs,ord,;
   ishod,recid_lpu,ispr MEMVAR 
  
  m.et       = merror.et
  m.ee       = merror.ee
  m.e_cod    = merror.e_cod
  m.e_ku     = merror.e_ku
  m.e_tip    = merror.e_tip
  m.err_mee  = merror.err_mee 
  m.osn230   = merror.osn230
  m.koeff    = merror.koeff
  m.e_period = merror.e_period
  m.straf    = merror.straf
  m.s_1      = merror.s_1
  m.s_2      = merror.s_2
   
  m.e_sall = 0
  m.s_lech = m.s_lech + m.s_all
  DO CASE 
   CASE m.TipAcc == 0 && Сводный счет
   CASE m.TipAcc == 1 && Амбулаторный счет
    m.dat1 = MIN(m.d_u, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
   CASE m.TipAcc == 2 && Дневной стационар
    m.dat1 = MIN(m.d_u-m.k_u, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
    m.dslast = m.dslast + k_u
   CASE m.TipAcc == 3 && Стационар
    m.dat1 = MIN(m.d_u-m.k_u+1, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
    m.dslast = m.dslast + k_u
  ENDCASE 
  IF m.et==m.TipOfExp
*   MESSAGEBOX(TRANSFORM(m.s_all,'999999.99'),0+64,m.et)
   DO CASE 
    CASE EMPTY(merror.err_mee)
    CASE LEFT(merror.err_mee,2)=='W0'
     m.TipOfMee = IIF(m.TipOfMee<=1, 1, m.TipOfMee)
    OTHERWISE 
     m.TipOfMee = 2 && Есть ошибки
     IF (!EMPTY(m.e_cod) AND m.cod != m.e_cod) OR (!EMPTY(m.e_ku) AND m.k_u != m.e_ku) OR ;
      (!EMPTY(m.e_tip) AND m.e_tip != m.tip) OR !EMPTY(m.koeff)
      IF m.koeff<=0
       m.e_sall = fsumm(m.e_cod, m.e_tip, m.e_ku, m.IsVed)
      ELSE 
       m.e_sall = ROUND(m.s_all * m.koeff,2)
      ENDIF 
     ENDIF 
   ENDCASE 
  ENDIF 
  INSERT INTO ttalon FROM MEMVAR 
 ENDSCAN 
 m.dat1 = IIF(m.TipAcc==3, m.dat1, m.dat1)
 
 SELECT people
  
 WAIT "ЗАПУСК WORD..." WINDOW NOWAIT 
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  oWord = CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 

 DO CASE 
  CASE m.TipAcc == 0
   m.AddToName = 'св'
   m.tipakt = 'свод'
  CASE m.TipAcc == 1
   m.AddToName = 'амб'
   m.tipakt = 'амб'
  CASE m.TipAcc == 2
   m.AddToName = 'дст'
   m.tipakt = 'дн/стац'
  CASE m.TipAcc == 3
   m.AddToName = 'ст'
   m.tipakt = 'стац'
 ENDCASE
 
 ooal = ALIAS()
 SELECT recid FROM ssacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) AND ;
  tipacc=m.tipacc AND sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) AND IsOk=IIF(m.TipOfMee=1, .t., .f.) INTO CURSOR rqwest NOCONSOLE 
 m.nfileid = recid
 USE 
 SELECT (ooal)
 
 IF m.nfileid>0
*   DocName = pmee+IIF(m.TipOfMee == 1,'\ActSSPlOK_','\ActSSPl_')+m.mcod+'_'+m.ppolis+'_'+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)+m.AddToName
  DocName = pmee+'\ssacts\'+PADL(m.nfileid,6,'0')
 ELSE 
  INSERT INTO ssacts (period,mcod,codexp,tipacc,isok,sn_pol,fam,im,ot) ;
   VALUES ;
  (m.gcperiod,m.mcod,INT(VAL(m.tipofexp)),m.tipacc,IIF(m.TipOfMee=1, .t., .f.),;
   PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot)
  m.nfileid = GETAUTOINCVALUE()
  DocName = pmee+'\ssacts\'+PADL(m.nfileid,6,'0')
  UPDATE ssacts SET actname=PADL(m.nfileid,6,'0')+'.doc', actdate=DATETIME() WHERE recid = m.nfileid
 ENDIF 
 
 IF fso.FileExists(DocName+'.doc')
  oFile = fso.GetFile(DocName+'.doc')
  DateCreated      = TTOC(oFile.DateCreated)
  DateLastAccessed = TTOC(oFile.DateLastAccessed)
  DateLastModified = TTOC(oFile.DateLastModified)
  RELEASE oFile
  
  IF MESSAGEBOX('ПО ВЫБРАННОМУ СЧЕТУ АКТ УЖЕ ФОРМИРОВАЛСЯ!'+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА СОЗДАНИЯ АКТА            : '+m.DateCreated+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ОТКРЫТИЯ АКТА : '+m.DateLastAccessed+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ИЗМЕНЕНИЯ АКТА: '+m.DateLastModified+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ВЫ ХОТИТЕ ПЕРЕФОРМИРОВАТЬ АКТ?',4+32,'') == 7 

   oWord.Quit
   USE IN ttalon
   CloseFiles()
   SELECT (oal)
   GO (orecno)
   RETURN
  ELSE 

   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
  
  ENDIF 

 ENDIF 

 oDoc = oWord.Documents.Add(pTempl+'\'+DotName)

 SELECT ttalon
 
* BROWSE 
 
 m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
 m.d_akt = DTOC(DATE())
 m.n_schet = STR(tYear,4)+PADL(tMonth,2,'0')
 IF m.TipAcc == 0
  m.dat1 = IIF(SEEK(m.sn_pol, 'people'), people.d_beg, {})
  m.dat2 = IIF(SEEK(m.sn_pol, 'people'), people.d_end, {})
 ENDIF 
 m.dslast = IIF(!INLIST(m.TipAcc,2,3), IIF(m.dat2-m.dat1>0, m.dat2-m.dat1, 1), m.dslast)
 m.ds = ds
 m.dsnam = IIF(SEEK(m.ds, 'mkb10'), ALLTRIM(mkb10.name_ds), '')
 m.pcod = pcod
 m.docfam = IIF(SEEK(m.pcod, 'doctor'), ALLTRIM(doctor.Fam)+' '+ALLTRIM(doctor.Im)+' '+ALLTRIM(doctor.Ot), '')
 m.fioexp = m.usrfam+' '+m.usrim+' '+m.usrot
 m.fioexp2 = LEFT(m.usrim,1)+'.'+LEFT(m.usrot,1)+'.'+m.usrfam


 oDoc.Bookmarks('tipakt').Select  
 oWord.Selection.TypeText(m.tipakt)
 oDoc.Bookmarks('tipofexp').Select  
 oWord.Selection.TypeText(m.ctipofexp)
 oDoc.Bookmarks('d_exp').Select  
 oWord.Selection.TypeText(DTOC(DATE()))
 oDoc.Bookmarks('n_akt').Select  
 oWord.Selection.TypeText(m.n_akt)
 oDoc.Bookmarks('d_akt').Select  
 oWord.Selection.TypeText(m.d_akt)
 oDoc.Bookmarks('fioexp').Select  
 oWord.Selection.TypeText(m.fioexp)
 oDoc.Bookmarks('d_akt2').Select  
 oWord.Selection.TypeText(m.d_akt)
 oDoc.Bookmarks('smo_name').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('lpu_name').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('sn_pol').Select  
 oWord.Selection.TypeText(ALLTRIM(m.sn_pol))
 oDoc.Bookmarks('n_schet').Select  
 oWord.Selection.TypeText(m.n_schet)
 oDoc.Bookmarks('c_i').Select  
 oWord.Selection.TypeText(ALLTRIM(c_i))
 oDoc.Bookmarks('ds').Select  
 oWord.Selection.TypeText(m.ds+', '+m.dsnam)
 oDoc.Bookmarks('dat1').Select  
 oWord.Selection.TypeText(DTOC(m.dat1))
 oDoc.Bookmarks('dat2').Select  
 oWord.Selection.TypeText(DTOC(m.dat2))
 oDoc.Bookmarks('docfam').Select  
 oWord.Selection.TypeText(ALLTRIM(m.docfam))
 oDoc.Bookmarks('s_lech').Select  
 oWord.Selection.TypeText(TRANSFORM(m.s_lech, '9999999.99'))
 oDoc.Bookmarks('dslast').Select  
 oWord.Selection.TypeText(TRANSFORM(m.dslast, '999'))
 oDoc.Bookmarks('fioexp2').Select  
 oWord.Selection.TypeText(m.fioexp2)

 IF m.TipOfMee == 1 && Все хорошо!

  nRow = 2
  m.tot_badsum = 0
  m.tot_goodsum = 0
  m.tot_straf = 0
  SCAN 
   oDoc.Tables(2).Cell(nRow,1).Select  && Код услуги
   oWord.Selection.InsertRows
   oWord.Selection.TypeText(STR(cod,6))
   oDoc.Tables(2).Cell(nRow,2).Select
   oWord.Selection.TypeText(DTOC(d_u))
   oDoc.Tables(2).Cell(nRow,3).Select

    m.tot_goodsum = m.tot_goodsum + s_all

    oDoc.Tables(2).Cell(nRow,5).Select
    oWord.Selection.TypeText(TRANSFORM(0, '999999.99'))
    oDoc.Tables(2).Cell(nRow,6).Select
    oWord.Selection.TypeText(TRANSFORM(0, '9999.99'))
    oDoc.Tables(2).Cell(nRow,7).Select
    oWord.Selection.TypeText(TRANSFORM(s_all, '999999.99'))
    
   nRow = nRow + 1
  ENDSCAN 
  USE 

  oDoc.Bookmarks('tot_badsum').Select  
  oWord.Selection.TypeText(TRANSFORM(m.tot_badsum, '999999.99'))
  oDoc.Bookmarks('tot_goodsum').Select  
  oWord.Selection.TypeText(TRANSFORM(m.tot_goodsum, '999999.99'))

  m.koplate   = m.tot_goodsum
  m.nekoplate = m.tot_badsum

  oDoc.Bookmarks('koplate').Select  
  oWord.Selection.TypeText(cpr(INT(m.koplate))+PADL(INT((m.koplate-int(m.koplate))*100),2,'0')+' КОП.')

  oDoc.Bookmarks('nekoplate').Select  
  oWord.Selection.TypeText(cpr(INT(m.nekoplate))+PADL(INT((m.nekoplate-INT(m.nekoplate))*100),2,'0')+' КОП.')
    
  oDoc.Bookmarks('resume').Select  
  oWord.Selection.TypeText('По представленному счету замечаний нет.')
  oDoc.Bookmarks('vivod').Select  
  oWord.Selection.TypeText('Счет подлежит оплате в полном объеме из средств ОМС.')

 ELSE && Есть ошибки!
 
  nRow = 2
  m.tot_badsum = 0
  m.tot_straf = 0
  m.tot_goodsum = 0
  SCAN 
   m.er_c = err_mee
   m.osn230 = IIF(SEEK(LEFT(UPPER(m.er_c),2), 'sookod'), sookod.osn230, '')	

   oDoc.Tables(2).Cell(nRow,1).Select  && Код услуги
   oWord.Selection.InsertRows
   oWord.Selection.TypeText(STR(cod,6))
   oDoc.Tables(2).Cell(nRow,2).Select
   oWord.Selection.TypeText(DTOC(d_u))
   oDoc.Tables(2).Cell(nRow,3).Select
   oWord.Selection.TypeText(m.osn230)
   oDoc.Tables(2).Cell(nRow,4).Select
   oWord.Selection.TypeText(m.er_c)

   IF !EMPTY(err_mee) AND UPPER(err_mee) != 'W0' AND et==m.TipOfExp
   
    m.tot_badsum = m.tot_badsum + IIF(koeff<=0, s_all-e_sall, e_sall)
    m.tot_goodsum = m.tot_goodsum + IIF(koeff<=0, e_sall, s_all-e_sall)
    m.tot_straf = m.tot_straf + straf*m.ynorm

    oDoc.Tables(2).Cell(nRow,5).Select
    oWord.Selection.TypeText(TRANSFORM(IIF(koeff<=0, s_all-e_sall, e_sall), '999999.99'))
    oDoc.Tables(2).Cell(nRow,6).Select
    oWord.Selection.TypeText(TRANSFORM(straf*m.ynorm, '9999.99'))
    oDoc.Tables(2).Cell(nRow,7).Select
    oWord.Selection.TypeText(TRANSFORM(IIF(koeff<=0, e_sall, s_all-e_sall), '999999.99'))
    
   ELSE && Если ошибки нет
    m.tot_goodsum = m.tot_goodsum + s_all

    oDoc.Tables(2).Cell(nRow,5).Select
    oWord.Selection.TypeText(TRANSFORM(0, '999999.99'))
    oDoc.Tables(2).Cell(nRow,6).Select
    oWord.Selection.TypeText(TRANSFORM(0, '9999.99'))
    oDoc.Tables(2).Cell(nRow,7).Select
    oWord.Selection.TypeText(TRANSFORM(s_all, '999999.99'))
    
   ENDIF 
   nRow = nRow + 1
  ENDSCAN 
  USE 
  oDoc.Bookmarks('tot_badsum').Select  
  oWord.Selection.TypeText(TRANSFORM(m.tot_badsum, '999999.99'))
  oDoc.Bookmarks('tot_straf').Select  
  oWord.Selection.TypeText(TRANSFORM(m.tot_straf, '999999.99'))
  oDoc.Bookmarks('tot_goodsum').Select  
  oWord.Selection.TypeText(TRANSFORM(m.tot_goodsum, '999999.99'))
     
  m.koplate   = m.s_lech - m.tot_badsum
  m.nekoplate = m.tot_badsum

  oDoc.Bookmarks('koplate').Select  
  oWord.Selection.TypeText(cpr(INT(m.koplate))+PADL(INT((m.koplate-int(m.koplate))*100),2,'0')+' КОП.')

  oDoc.Bookmarks('nekoplate').Select  
  oWord.Selection.TypeText(cpr(INT(m.nekoplate))+PADL(INT((m.nekoplate-INT(m.nekoplate))*100),2,'0')+' КОП.')
 ENDIF 
 oDoc.SaveAs(DocName,0)

 =CloseFiles()
 SELECT (oal)
 GO (orecno)

 WAIT CLEAR 
 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 

RETURN 

FUNCTION OpenFiles(ppath)
 IF m.lWasUsedSprLpu = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpu', 'sprlpu', 'shared', 'mcod')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedTarif = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shared', 'cod')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedMkb = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\mkb10', 'mkb10', 'shared', 'ds')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedSooKod = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'shar', 'er_c')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF
 ENDIF 

 IF m.lWasUsedTalon = .F.
  IF OpenFile(pPath+'\Talon', 'Talon', 'SHARED', 'sn_pol')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF
 ENDIF  
 
 IF m.lWasUsedPeople = .F.
  IF OpenFile(pPath+'\people', 'people', 'SHARED', 'sn_pol')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 
 
 IF m.lWasUsedDoctor = .F.
  IF OpenFile(pPath+'\doctor', 'doctor', 'SHARED', 'pcod')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 
  
 IF m.lWasUsedSSActs = .F.
  IF OpenFile(pmee+'\ssacts\ssacts', 'ssacts', 'SHARED')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedMError = .F.
  IF OpenFile(pPath+'\m'+m.mcod, 'merror', 'SHARED', 'id_et')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

RETURN .T.

FUNCTION CloseFiles
 IF m.lWasUsedSprLpu = .F.
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
 ENDIF 
 IF m.lWasUsedTarif = .F.
  IF USED('tarif')
   USE IN tarif 
  ENDIF 
 ENDIF 
 IF m.lWasUsedMkb = .F.
  IF USED('mkb10')
   USE IN mkb10
  ENDIF 
 ENDIF 
 IF m.lWasUsedSooKod = .F.
  IF USED('sookod')
   USE IN sookod
  ENDIF 
 ENDIF 
 IF m.lWasUsedTalon = .F.
  IF USED('talon')
   USE IN talon 
  ENDIF 
 ENDIF 
 IF m.lWasUsedPeople = .F.
  IF USED('people')
   USE IN people
  ENDIF 
 ENDIF 
 IF m.lWasUsedDoctor = .F.
  IF USED('doctor')
   USE IN doctor
  ENDIF 
 ENDIF 
 IF m.lWasUsedSSActs = .F.
  IF USED('ssacts')
   USE IN ssacts
  ENDIF 
 ENDIF 
 IF m.lWasUsedMError = .F.
  IF USED('merror')
   USE IN merror
  ENDIF 
 ENDIF 
RETURN 
