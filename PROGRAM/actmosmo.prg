PROCEDURE  actmosmo(lcPath, cmcod, IsVisible, IsQuit)
 IF !fso.FileExists(ptempl+'\svqqmmyy.dot')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ SVQQMMYY.DOT!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.lpuname = IIF(SEEK(cmcod, 'sprlpu', 'mcod'), ALLTRIM(sprlpu.fullname), '')
 m.cokr = IIF(SEEK(cmcod, 'sprlpu', 'mcod'), sprlpu.cokr, '00')
 m.cokrname = ''
 DO CASE 
  CASE m.cokr = '00'
   m.cokrname = 'Не определен'
  CASE m.cokr = '01'
   m.cokrname = 'ЦАО'
  CASE m.cokr = '02'
   m.cokrname = 'САО'
  CASE m.cokr = '03'
   m.cokrname = 'СВАО'
  CASE m.cokr = '04'
   m.cokrname = 'ВАО'
  CASE m.cokr = '05'
   m.cokrname = 'ЮВАО'
  CASE m.cokr = '06'
   m.cokrname = 'ЮАО'
  CASE m.cokr = '07'
   m.cokrname = 'ЮЗАО'
  CASE m.cokr = '08'
   m.cokrname = 'ЗАО'
  CASE m.cokr = '09'
   m.cokrname = 'СЗАО'
  CASE m.cokr = '10'
   m.cokrname = 'ЗелАО'
  CASE m.cokr = '11'
   m.cokrname = 'Вне пределов г. Москвы'
  CASE m.cokr = '12'
   m.cokrname = 'Новомосковский АО'
  CASE m.cokr = '13'
   m.cokrname = 'Троицкий АО'
 ENDCASE 
 m.lpuname = IIF(SEEK(cmcod, 'sprlpu', 'mcod'), ALLTRIM(sprlpu.fullname), '')+' '+;
  ' ('+cmcod+'), '+m.cokrname+'('+m.cokr+')'
* m.IsPilot = .f.
* IF fso.FileExists(pcommon+'\pilot.dbf')
*  IF OpenFile(pcommon+'\pilot', 'pilot', 'shar', 'mcod')<=0
*   m.IsPilot = IIF(SEEK(cmcod, 'pilot'), .t., .f.)
*   USE IN pilot
*  ENDIF 
* ENDIF 
 
 IF !fso.FolderExists(lcPath)
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ'+CHR(13)+CHR(10)+;
   UPPER(ALLTRIM(lcPath))+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF !fso.FileExists(lcPath+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ'+CHR(13)+CHR(10)+;
   UPPER(ALLTRIM(lcPath))+'\PEOPLE.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcPath+'\etrl'+m.qcod+'.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ'+CHR(13)+CHR(10)+;
   UPPER(ALLTRIM(lcPath))+'\ETRL'+UPPER(m.qcod)+'.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(lcPath+'\people', 'people', 'shar')>0
  IF USED('people')
   USE IN people 
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(lcPath+'\etrl'+m.qcod, 'etrl', 'shar')>0
  USE IN people
  IF USED('etrl')
   USE IN etrl 
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT people
 m.totrecs       = RECCOUNT('people')
 m.recsattaches  = 0
 m.recsattaches2 = 0
 m.toterrs       = 0
 m.errsdoubles   = 0
 m.totaccepted   = 0
 m.totaccepted2  = 0

 SCAN 
  m.recsattaches  = m.recsattaches  + IIF(EMPTY(date_out),1,0)
  m.recsattaches2 = m.recsattaches2 + IIF(EMPTY(date_out) and spos='2', 1, 0)
  m.toterrs       = m.toterrs + IIF(!EMPTY(c_err) AND c_err!='OK',1, 0)
  m.errsdoubles   = m.errsdoubles + IIF(INLIST(c_err,'F8','F9'),1, 0)
  m.totaccepted   = m.totaccepted + IIF(c_err='OK', 1, 0)
  m.totaccepted2  = m.totaccepted2 + IIF(c_err='OK' and spos='2', 1, 0)
  
  m.vozr = ROUND((m.tdat1 - dr)/365.25,2)
  m.w    = w 

  IF c_err='OK'
*   m.totchildren  = m.totchildren  + IIF(m.vozr<5,1,0)
*   m.malechildren = m.malechildren + IIF(m.vozr<5 and m.w=1,1,0)
*   m.femchildren  = m.femchildren  + IIF(m.vozr<5 and m.w=2,1,0)
*   m.totteens     = m.totteens     + IIF(BETWEEN(m.vozr,5,17.99),1,0)
*   m.maleteens    = m.maleteens    + IIF(BETWEEN(m.vozr,5,17.99) and m.w=1,1,0)
*   m.femteens     = m.femteens     + IIF(BETWEEN(m.vozr,5,17.99) and m.w=2,1,0)
*   m.totmen       = m.totmen       + IIF(BETWEEN(m.vozr,18,60) AND m.w=1,1,0)
*   m.totwomen     = m.totwomen     + IIF(BETWEEN(m.vozr,18,55) AND m.w=2,1,0)
*   m.totpensmen   = m.totpensmen   + IIF(m.vozr>60 and m.w=1,1,0)
*   m.totpenswomen = m.totpenswomen + IIF(m.vozr>55 and m.w=2,1,0)
  ENDIF 

 ENDSCAN 

* MESSAGEBOX(TRANSFORM(m.totaccepted,'999999.99'),0+64,'')
 m.totaccepted = m.recsattaches -  RECCOUNT('etrl')
 USE IN etrl 
 USE IN people
 SELECT aisoms
 
 
* IF m.IsPilot
*   m.totchildren  = ch04m + ch04f
*   m.malechildren = ch04m 
*   m.femchildren  = ch04f 
*  m.totteens     = ch517m + ch517f
*   m.maleteens    = ch517m
*   m.femteens     = ch517f
*   m.totmen       = m1859
*   m.totwomen     = f1854
*   m.totpensmen   = m60
*   m.totpenswomen = f55
* ENDIF 

 DotName = pTempl + '\svqqmmyy.dot'
 DocName = lcPath + '\sv' + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)

 TRY 
 oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 oDoc = oWord.Documents.Add(dotname)

 oDoc.Bookmarks('period').Select  
* oWord.Selection.TypeText(NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+' '+SUBSTR(m.gcperiod,1,4)+' года')
 IF m.tMonth < 12
  m.period = '01 '+NameOfMonth2(tMonth+1)+' '+STR(tYear,4)+' года'
 ELSE 
  m.period = '01 января '+STR(tYear+1,4)+' года'
 ENDIF 
 oWord.Selection.TypeText(m.period)
* oWord.Selection.TypeText('01 '+NameOfMonth2(VAL(SUBSTR(m.gcperiod,5,2))+1)+' '+SUBSTR(m.gcperiod,1,4)+' года')
* oWord.Selection.TypeText(NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2))+1)+' '+SUBSTR(m.gcperiod,1,4)+' года')
 oDoc.Bookmarks('lpuname').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('smoname').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('ggggmm').Select  
 oWord.Selection.TypeText(STR(tYear,4)+PADL(tMonth,2,'0'))
 
 oDoc.Bookmarks('period2').Select  
* oWord.Selection.TypeText(' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+' '+SUBSTR(m.gcperiod,1,4)+' года')
* oWord.Selection.TypeText('01 '+NameOfMonth2(VAL(SUBSTR(m.gcperiod,5,2))+1)+' '+SUBSTR(m.gcperiod,1,4)+' года')
 IF m.tMonth < 12
  m.period2 = ' '+NameOfMonth(tMonth)+' '+STR(tYear,4)+' года'
 ELSE 
  m.period2 = ' январь '+STR(tYear+1,4)+' года'
 ENDIF 
* oWord.Selection.TypeText(NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2))+1)+' '+SUBSTR(m.gcperiod,1,4)+' года')
 oWord.Selection.TypeText(m.period2)
 oDoc.Bookmarks('lpuname2').Select  
 oWord.Selection.TypeText(' '+m.lpuname)
 oDoc.Bookmarks('smoname2').Select  
 oWord.Selection.TypeText(' '+m.qname)

 WITH oDoc.Tables(1)
  .Cell(4,1).Select
  oWord.Selection.TypeText(TRANSFORM(ch01m+ch01f+ch14m+ch14f+ch514m+ch514f+ch1517m+ch1517f,'99999'))
  .Cell(4,2).Select
  oWord.Selection.TypeText(TRANSFORM(ch01m,'99999'))
  .Cell(4,3).Select
  oWord.Selection.TypeText(TRANSFORM(ch01f,'99999'))
  .Cell(4,4).Select
  oWord.Selection.TypeText(TRANSFORM(ch14m,'99999'))
  .Cell(4,5).Select
  oWord.Selection.TypeText(TRANSFORM(ch14f,'99999'))
  .Cell(4,6).Select
  oWord.Selection.TypeText(TRANSFORM(ch514m,'99999'))
  .Cell(4,7).Select
  oWord.Selection.TypeText(TRANSFORM(ch514f,'99999'))
  .Cell(4,8).Select
  oWord.Selection.TypeText(TRANSFORM(ch1517m,'99999'))
  .Cell(4,9).Select
  oWord.Selection.TypeText(TRANSFORM(ch1517f,'99999'))
 ENDWITH 

 WITH oDoc.Tables(2)
  .Cell(4,1).Select
  oWord.Selection.TypeText(TRANSFORM(m1824+f1824+m2534+f2534+m3544+f3544+m4559+f4559,'99999'))
  .Cell(4,2).Select
  oWord.Selection.TypeText(TRANSFORM(m1824,'99999'))
  .Cell(4,3).Select
  oWord.Selection.TypeText(TRANSFORM(f1824,'99999'))
  .Cell(4,4).Select
  oWord.Selection.TypeText(TRANSFORM(m2534,'99999'))
  .Cell(4,5).Select
  oWord.Selection.TypeText(TRANSFORM(f2534,'99999'))
  .Cell(4,6).Select
  oWord.Selection.TypeText(TRANSFORM(m3544,'99999'))
  .Cell(4,7).Select
  oWord.Selection.TypeText(TRANSFORM(f3544,'99999'))
  .Cell(4,8).Select
  oWord.Selection.TypeText(TRANSFORM(m4559,'99999'))
  .Cell(4,9).Select
  oWord.Selection.TypeText(TRANSFORM(f4559,'99999'))
 ENDWITH 

 WITH oDoc.Tables(3)
  .Cell(4,1).Select
  oWord.Selection.TypeText(TRANSFORM(m6068+f5564+m69+f65,'99999'))
  .Cell(4,2).Select
  oWord.Selection.TypeText(TRANSFORM(m6068,'99999'))
  .Cell(4,3).Select
  oWord.Selection.TypeText(TRANSFORM(f5564,'99999'))
  .Cell(4,4).Select
  oWord.Selection.TypeText(TRANSFORM(m69,'99999'))
  .Cell(4,5).Select
  oWord.Selection.TypeText(TRANSFORM(f65,'99999'))

  m.tpaz = ch01m+ch01f+ch14m+ch14f+ch514m+ch514f+ch1517m+ch1517f+;
   m1824+f1824+m2534+f2534+m3544+f3544+m4559+f4559+;
   m6068+f5564+m69+f65
   
  .Cell(4,6).Select
  oWord.Selection.TypeText(TRANSFORM(m.tpaz,'99999'))
 ENDWITH 

* oDoc.Bookmarks('totrecs').Select  
* oWord.Selection.TypeText(ALLTRIM(STR(m.totrecs)))
 oDoc.Bookmarks('recsattaches').Select  
 oWord.Selection.TypeText(ALLTRIM(STR(m.recsattaches)))
* oDoc.Bookmarks('recsattaches2').Select  
* oWord.Selection.TypeText(ALLTRIM(STR(m.recsattaches2)))
* oDoc.Bookmarks('toterrs').Select  
* oWord.Selection.TypeText(ALLTRIM(STR(m.toterrs)))
* oDoc.Bookmarks('errsdoubles').Select  
* oWord.Selection.TypeText(ALLTRIM(STR(m.errsdoubles)))
 oDoc.Bookmarks('totaccepted').Select  
 oWord.Selection.TypeText(ALLTRIM(STR(m.totaccepted)))
* oDoc.Bookmarks('totaccepted2').Select  
* oWord.Selection.TypeText(ALLTRIM(STR(m.totaccepted2)))

 TRY 
  oDoc.SaveAs(DocName, 17)
 CATCH 
 ENDTRY 
 oDoc.SaveAs(DocName, 0)
 
 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 
 
RETURN 
 
