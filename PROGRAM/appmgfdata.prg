PROCEDURE AppMgfData
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
  
  m.ch_mgf = oExcel.Cells(nCell,7).Value
  m.ad_mgf = oExcel.Cells(nCell,16).Value + oExcel.Cells(nCell,25).Value
  
  m.ch01m   = oExcel.Cells(nCell,8).Value
  m.ch01f   = oExcel.Cells(nCell,9).Value
  m.ch14m   = oExcel.Cells(nCell,10).Value
  m.ch14f   = oExcel.Cells(nCell,11).Value
  m.ch514m  = oExcel.Cells(nCell,12).Value
  m.ch514f  = oExcel.Cells(nCell,13).Value
  m.ch1517m = oExcel.Cells(nCell,14).Value && ??
  m.ch1517f = oExcel.Cells(nCell,15).Value && ??
  m.m1824   = oExcel.Cells(nCell,17).Value
  m.f1824   = oExcel.Cells(nCell,18).Value
  m.m2534   = oExcel.Cells(nCell,19).Value
  m.f2534   = oExcel.Cells(nCell,20).Value
  m.m3544   = oExcel.Cells(nCell,21).Value
  m.f3544   = oExcel.Cells(nCell,22).Value
  m.m4559   = oExcel.Cells(nCell,23).Value
  m.f4559   = oExcel.Cells(nCell,24).Value
  m.m6068   = oExcel.Cells(nCell,26).Value
  m.f5564   = oExcel.Cells(nCell,27).Value
  m.m69     = oExcel.Cells(nCell,28).Value
  m.f65     = oExcel.Cells(nCell,29).Value
  
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
   CASE VARTYPE(m.ch01m)='C'
    m.ch01m = INT(VAL(m.ch01m))
   CASE VARTYPE(m.ch01m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch01f)='C'
    m.ch01f = INT(VAL(m.ch01f))
   CASE VARTYPE(m.ch01f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch14m)='C'
    m.ch14m = INT(VAL(m.ch14m))
   CASE VARTYPE(m.ch14m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch14f)='C'
    m.ch14f = INT(VAL(m.ch14f))
   CASE VARTYPE(m.ch14f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch514m)='C'
    m.ch514m = INT(VAL(m.ch514m))
   CASE VARTYPE(m.ch514m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch514f)='C'
    m.ch514f = INT(VAL(m.ch514f))
   CASE VARTYPE(m.ch514f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch1517m)='C'
    m.ch1517m = INT(VAL(m.ch1517m))
   CASE VARTYPE(m.ch1517m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch1517f)='C'
    m.ch1517f = INT(VAL(m.ch1517f))
   CASE VARTYPE(m.ch1517f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m1824)='C'
    m.m1824 = INT(VAL(m.m1824))
   CASE VARTYPE(m.m1824)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f1824)='C'
    m.f1824 = INT(VAL(m.f1824))
   CASE VARTYPE(m.f1824)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m2534)='C'
    m.m2534 = INT(VAL(m.m2534))
   CASE VARTYPE(m.m2534)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f2534)='C'
    m.f2534 = INT(VAL(m.f2534))
   CASE VARTYPE(m.f2534)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m3544)='C'
    m.m3544 = INT(VAL(m.m3544))
   CASE VARTYPE(m.m3544)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f3544)='C'
    m.f3544 = INT(VAL(m.f3544))
   CASE VARTYPE(m.f3544)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m4559)='C'
    m.m4559 = INT(VAL(m.m4559))
   CASE VARTYPE(m.m4559)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f4559)='C'
    m.f4559 = INT(VAL(m.f4559))
   CASE VARTYPE(m.f4559)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m6068)='C'
    m.m6068 = INT(VAL(m.m6068))
   CASE VARTYPE(m.m6068)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f5564)='C'
    m.f5564 = INT(VAL(m.f5564))
   CASE VARTYPE(m.f5564)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m69)='C'
    m.m69 = INT(VAL(m.m69))
   CASE VARTYPE(m.m69)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f65)='C'
    m.f65 = INT(VAL(m.f65))
   CASE VARTYPE(m.f65)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 


  IF SEEK(m.nlpuid, 'aisoms')

   UPDATE aisoms SET ch_mgf=m.ch_mgf, ad_mgf=m.ad_mgf, ch01m=m.ch01m, ch01f=m.ch01f,;
    ch14m=m.ch14m, ch14f=m.ch14f, ;
    ch514m=m.ch514m, ch514f=m.ch514f, ch1517m=m.ch1517m, ch1517f=m.ch1517f, ;
    m1824=m.m1824, f1824=m.f1824, m2534=m.m2534, f2534=m.f2534, m3544=m.m3544, f3544=m.f3544, ;
    m4559=m.m4559, f4559=m.f4559, m6068=m.m6068, f5564=m.f5564, m69=m.m69, f65=m.f65 ;
    WHERE lpuid=m.nlpuid

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