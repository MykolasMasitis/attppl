PROCEDURE AppMgfDataO
 IF MESSAGEBOX(CHR(13)+CHR(10)+'«¿√–”«»“‹ ◊»—À≈ÕÕŒ—“‹ œ–» –≈œÀ≈Õ»ﬂ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 m.mmyy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)
 flname = 'OMS'+m.qcod+m.mmyy+'.xls'
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+flname)
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ Œ¡Õ¿–”∆≈Õ ‘¿…À '+flname+CHR(13)+CHR(10)+;
   '¬ ƒ»–≈ “Œ–»» '+pbase+'\'+gcperiod+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF'+CHR(13)+CHR(10)+;
   '¬ ƒ»–≈ “Œ–»» '+pbase+'\'+gcperiod+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 WAIT "«¿œ”—  EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 
 m.IsVisible = .f. 
 m.IsQuit    = .t.
 oDoc = oExcel.WorkBooks.Open(pbase+'\'+gcperiod+'\'+flname,.T.)
 
 oExcel.Sheets(1).Select
 
 m.pnLpu = 0
 nCells=200
 FOR nCell=1 TO nCells
  m.nLpu  = oExcel.Cells(nCell,1).Value
  IF EMPTY(m.nLpu)
   LOOP 
  ENDIF 
  DO CASE 
   CASE VARTYPE(m.nLpu)='C'
    m.nLpu = INT(VAL(m.nLpu))
   CASE VARTYPE(m.nLpu)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  
  IF m.nLpu=0
   LOOP 
  ENDIF 
  
  IF m.nLpu-m.pnLpu != 1
   EXIT 
  ENDIF 
  m.nlpuid = oExcel.Cells(nCell,2).Value
  DO CASE 
   CASE VARTYPE(m.nlpuid)='C'
    m.nlpuid = INT(VAL(m.nlpuid))
   CASE VARTYPE(m.nlpuid)='N'
   OTHERWISE 
    EXIT 
  ENDCASE 
  
  m.ch04m  = oExcel.Cells(nCell,8).Value
  m.ch04f  = oExcel.Cells(nCell,9).Value
  m.ch517m = oExcel.Cells(nCell,10).Value
  m.ch517f = oExcel.Cells(nCell,11).Value
  m.m1859  = oExcel.Cells(nCell,13).Value
  m.f1854  = oExcel.Cells(nCell,14).Value
  m.m60    = oExcel.Cells(nCell,16).Value
  m.f55    = oExcel.Cells(nCell,17).Value
  
  m.ch_mgf = oExcel.Cells(nCell,7).Value
  m.ad_mgf = oExcel.Cells(nCell,12).Value + oExcel.Cells(nCell,15).Value
  
  DO CASE 
   CASE VARTYPE(m.ch_mgf)='C'
    m.ch_mgf = INT(VAL(m.ch_mgf))
   CASE VARTYPE(m.ch_mgf)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.ad_mgf)='C'
    m.ad_mgf = INT(VAL(m.ad_mgf))
   CASE VARTYPE(m.ad_mgf)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.ch04m)='C'
    m.ch04m = INT(VAL(m.ch04m))
   CASE VARTYPE(m.ch04m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.ch04f)='C'
    m.ch04f = INT(VAL(m.ch04f))
   CASE VARTYPE(m.ch04f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.ch517m)='C'
    m.ch517m = INT(VAL(m.ch517m))
   CASE VARTYPE(m.ch517m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.ch517f)='C'
    m.ch517f = INT(VAL(m.ch517f))
   CASE VARTYPE(m.ch517f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.m1859)='C'
    m.m1859 = INT(VAL(m.m1859))
   CASE VARTYPE(m.m1859)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.f1854)='C'
    m.f1854 = INT(VAL(m.f1854))
   CASE VARTYPE(m.f1854)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.m60)='C'
    m.m60 = INT(VAL(m.m60))
   CASE VARTYPE(m.m60)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.f55)='C'
    m.f55 = INT(VAL(m.f55))
   CASE VARTYPE(m.f55)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  IF SEEK(m.nlpuid, 'aisoms')
   UPDATE aisoms SET ch_mgf=m.ch_mgf, ad_mgf=m.ad_mgf, ;
    ch04m=m.ch04m, ch04f=m.ch04f, ch517m=m.ch517m, ch517f=m.ch517f, ;
    m1859=m.m1859, f1854=m.f1854, m60=m.m60, f55=m.f55 WHERE lpuid=m.nlpuid
  ENDIF 
  
  IF m.nLpu>=1
   m.pnLpu = m.nLpu
  ENDIF 
  
 NEXT 

 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

 WAIT CLEAR 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
RETURN 