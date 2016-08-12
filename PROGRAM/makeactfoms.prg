PROCEDURE MakeActFoms
 IF MESSAGEBOX(CHR(13)+CHR(10)+'—‘Œ–Ã»–Œ¬¿“‹ —¬ŒƒÕ€… ¿ “?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\sv.dot')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À ÿ¿¡ÀŒÕ¿ '+'SV.DOT'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'—¬ŒƒÕ€… –≈√»—“– Õ≈ —Œ¡–¿Õ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'ais', 'shar')>0
  IF USED('ais')
   USE IN ais
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ELSE 
  SELECT * FROM ais ORDER BY cokr, mcod INTO CURSOR curais
  USE IN ais
 ENDIF 

 dotname = ptempl+'\sv.dot'
 docname = pbase+'\'+gcperiod+'\vv'+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)
 
 WAIT "«¿œ”— ¿ MS WORD..." WINDOW NOWAIT 
 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 

 oDoc   = oWord.Documents.Add(dotname)

 oDoc.Bookmarks('period').Select  
 oWord.Selection.TypeText(+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+' '+SUBSTR(m.gcperiod,1,4)+' „Ó‰‡')
 oDoc.Bookmarks('smoname').Select  
 oWord.Selection.TypeText(m.qname)

 oTable = oDoc.Tables(1)

 m._totrecs       = 0
 m._recsattaches  = 0
 m._recsattaches2 = 0
 m._toterrs       = 0
 m._errsdoubles   = 0
 m._totaccepted   = 0
 m._totaccepted2  = 0
 
 m._totchildren  = 0 
 m._malechildren = 0 
 m._femchildren  = 0
 m._totteens     = 0
 m._maleteens    = 0
 m._femteens     = 0
 m._totpeople    = 0
 m._totmen       = 0
 m._totwomen     = 0
 m._totpens      = 0
 m._totpensmen   = 0
 m._totpenswomen = 0

 SELECT curais
 nCell = 0
 SCAN 
  m.etval = pazin - pazbad
  m.mcod  = mcod
  m.lpuid = STR(lpuid,4)
  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.name), '')
  m.cokr    = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.cokr, '')
  IF EMPTY(bname)
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+gcperiod+'\'+mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+gcperiod+'\'+mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+gcperiod+'\'+mcod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people
    SELECT curais
   ENDIF 
   LOOP 
  ENDIF 
  IF RECCOUNT('people')<=0
   USE IN people
   SELECT curais
   LOOP 
  ENDIF 

  SELECT people

  m.totrecs       = RECCOUNT('people')
*  COUNT FOR EMPTY(date_out) TO m.totrecs
  m.recsattaches  = 0
  m.recsattaches2 = 0
  m.toterrs       = 0
  m.errsdoubles   = 0
  m.totaccepted   = 0
  m.totaccepted2  = 0
 
  m.totchildren  = 0 
  m.malechildren = 0 
  m.femchildren  = 0
  m.totteens     = 0
  m.maleteens    = 0
  m.femteens     = 0
  m.totpeople    = 0
  m.totmen       = 0
  m.totwomen     = 0
  m.totpens      = 0
  m.totpensmen   = 0
  m.totpenswomen = 0

  WAIT m.mcod+'...' WINDOW NOWAIT 
  SCAN 
   m.recsattaches  = m.recsattaches  + IIF(EMPTY(date_out),1,0)
   m.recsattaches2 = m.recsattaches2 + IIF(EMPTY(date_out) and spos='2', 1, 0)
   m.toterrs       = m.toterrs + IIF(!EMPTY(c_err) AND c_err!='OK',1, 0)
   m.errsdoubles   = m.errsdoubles + IIF(INLIST(c_err,'F8','F9'),1, 0)
   m.totaccepted   = m.totaccepted + IIF(EMPTY(date_out) and c_err='OK', 1, 0)
   m.totaccepted2  = m.totaccepted2 + IIF(EMPTY(date_out) and c_err='OK' and spos='2', 1, 0)
  
   m.vozr = ROUND((m.tdat1 - dr)/365.25,2)
   m.w    = w 

   IF c_err='OK'
    m.totchildren  = m.totchildren  + IIF(m.vozr<5,1,0)
    m.malechildren = m.malechildren + IIF(m.vozr<5 and m.w=1,1,0)
    m.femchildren  = m.femchildren  + IIF(m.vozr<5 and m.w=2,1,0)
    m.totteens     = m.totteens     + IIF(BETWEEN(m.vozr,5,17.99),1,0)
    m.maleteens    = m.maleteens    + IIF(BETWEEN(m.vozr,5,17.99) and m.w=1,1,0)
    m.femteens     = m.femteens     + IIF(BETWEEN(m.vozr,5,17.99) and m.w=2,1,0)
    m.totmen       = m.totmen       + IIF(BETWEEN(m.vozr,18,60) AND m.w=1,1,0)
    m.totwomen     = m.totwomen     + IIF(BETWEEN(m.vozr,18,55) AND m.w=2,1,0)
    m.totpensmen   = m.totpensmen   + IIF(m.vozr>60 and m.w=1,1,0)
    m.totpenswomen = m.totpenswomen + IIF(m.vozr>55 and m.w=2,1,0)
   ENDIF 
  
  ENDSCAN 
  USE 

  m._totrecs       = m._totrecs + m.totrecs
  m._recsattaches  = m._recsattaches + m.recsattaches
  m._recsattaches2 = m._recsattaches2 + m.recsattaches2
  m._toterrs       = m._toterrs + m.toterrs
  m._errsdoubles   = m._errsdoubles + m.errsdoubles
  m._totaccepted   = m._totaccepted + m.totaccepted
  m._totaccepted2  = m._totaccepted2 + m.totaccepted2
 
  m._totchildren  = m._totchildren + m.totchildren
  m._malechildren = m._malechildren + m.malechildren
  m._femchildren  = m._femchildren + m.femchildren
  m._totteens     = m._totteens + m.totteens
  m._maleteens    = m._maleteens + m.maleteens
  m._femteens     = m._femteens + m.femteens
  m._totpeople    = m._totpeople + m.totpeople
  m._totmen       = m._totmen + m.totmen
  m._totwomen     = m._totwomen + m.totwomen 
  m._totpens      = m._totpens + m.totpens
  m._totpensmen   = m._totpensmen + m.totpensmen
  m._totpenswomen = m._totpenswomen + m.totpenswomen
  
  IF m.totaccepted!= m.etval
*   MESSAGEBOX('m.totaccepted = '+TRANSFORM(m.totaccepted,'99999')+CHR(13)+CHR(10)+;
    'm.etval (pazin-pazbad) = '+TRANSFORM(m.etval,'99999')+CHR(13)+CHR(10),0+64,mcod)
  ENDIF 

  oTable.Cell(4+nCell,1).Select
  oWord.Selection.TypeText(m.lpuname)
  oTable.Cell(4+nCell,2).Select
  oWord.Selection.TypeText(m.lpuid)
  oTable.Cell(4+nCell,3).Select
  oWord.Selection.TypeText(m.cokr)
  oTable.Cell(4+nCell,4).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totrecs)))
  oTable.Cell(4+nCell,5).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totaccepted)))
  oTable.Cell(4+nCell,6).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totaccepted2)))

  oTable.Cell(4+nCell,7).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totchildren)))
  oTable.Cell(4+nCell,8).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.malechildren)))
  oTable.Cell(4+nCell,9).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.femchildren)))
  oTable.Cell(4+nCell,10).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totteens)))
  oTable.Cell(4+nCell,11).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.maleteens)))
  oTable.Cell(4+nCell,12).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.femteens)))
  oTable.Cell(4+nCell,13).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totmen+m.totwomen)))
  oTable.Cell(4+nCell,14).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totmen)))
  oTable.Cell(4+nCell,15).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totwomen)))
  oTable.Cell(4+nCell,16).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totpensmen+m.totpenswomen)))
  oTable.Cell(4+nCell,17).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totpensmen)))
  oTable.Cell(4+nCell,18).Select
  oWord.Selection.TypeText(ALLTRIM(STR(m.totpenswomen)))

  oTable.Cell(4+nCell,1).Select
  oWord.Selection.InsertRowsBelow

  nCell = nCell + 1

  WAIT CLEAR 
  SELECT curais 
  
 ENDSCAN 

 USE 
 USE IN sprlpu

* nCell = nCell - 1
 oTable.Cell(4+nCell,4).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totrecs)))
 oTable.Cell(4+nCell,5).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totaccepted)))
 oTable.Cell(4+nCell,6).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totaccepted2)))

 oTable.Cell(4+nCell,7).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totchildren)))
 oTable.Cell(4+nCell,8).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._malechildren)))
 oTable.Cell(4+nCell,9).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._femchildren)))
 oTable.Cell(4+nCell,10).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totteens)))
 oTable.Cell(4+nCell,11).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._maleteens)))
 oTable.Cell(4+nCell,12).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._femteens)))
 oTable.Cell(4+nCell,13).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totmen+m._totwomen)))
 oTable.Cell(4+nCell,14).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totmen)))
 oTable.Cell(4+nCell,15).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totwomen)))
 oTable.Cell(4+nCell,16).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totpensmen+m._totpenswomen)))
 oTable.Cell(4+nCell,17).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totpensmen)))
 oTable.Cell(4+nCell,18).Select
 oWord.Selection.TypeText(ALLTRIM(STR(m._totpenswomen)))

* oRange = oWord.Range(oTable.Cell(4+nCell-1,1), oTable.Cell(4+nCell-1,3))
* oRange.Merge
 
 WAIT "Œ—“¿ÕŒ¬ ¿ MS WORD..." WINDOW NOWAIT  
 TRY 
  oDoc.SaveAs(DocName, 17)
 CATCH 
 ENDTRY 
 oDoc.SaveAs(DocName, 0)
 oWord.Visible = .T.
* oDoc.Close
* oWord.Quit
 WAIT CLEAR 

RETURN 