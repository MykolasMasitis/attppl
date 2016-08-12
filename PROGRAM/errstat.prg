PROCEDURE ErrStat
 IF MESSAGEBOX('ÏÎÑ×ÈÒÀÒÜ ÑÒÀÒÈÑÒÈÊÓ ÏÎ ÎØÈÁÊÀÌ?',4+32,'')=7
  RETURN 
 ENDIF 
 
* CREATE CURSOR curstat (c_err c(2), m01 n(6), m02 n(6), m03 n(6), m04 n(6), m05 n(6), m06 n(6), ;
  m07 n(7), m08 n(6), m09 n(6), m10 n(6), m11 n(6), m12 n(6)) 
  
 DIMENSION dimstat(15,12)
 dimstat = 0 
 
 FOR nmon=1 TO tmonth
  m.ppath = STR(m.tyear,4)+PADL(m.nmon,2,'0')
  WAIT m.ppath WINDOW NOWAIT 
  =OneStat(nmon)
  WAIT CLEAR 
 ENDFOR 

 dotname = ptempl+'\errstat.xlt'
 docname = pout+'\errstat'

 IF !fso.FileExists(dotname)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ ERRSTAT.XLT'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF fso.FileExists(docname+'.xls')
  fso.DeleteFile(docname+'.xls')
 ENDIF 

 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 

 WITH oExcel
  .ReferenceStyle= -4150  && xlR1C1
  .SheetsInNewWorkbook = 1
 ENDWITH 
 oDoc = oExcel.WorkBooks.Add(dotname)
 
 WITH oExcel
  
  FOR m.iii = 1 TO 6
   .Cells(03,03+m.iii).Value  = dimstat(1,m.iii)
   .Cells(04,03+m.iii).Value  = dimstat(2,m.iii)
   .Cells(05,03+m.iii).Value  = dimstat(3,m.iii)
   .Cells(06,03+m.iii).Value  = dimstat(4,m.iii)
   .Cells(07,03+m.iii).Value  = dimstat(5,m.iii)
   .Cells(08,03+m.iii).Value  = dimstat(6,m.iii)
   .Cells(09,03+m.iii).Value  = dimstat(7,m.iii)
   .Cells(10,03+m.iii).Value  = dimstat(8,m.iii)
   .Cells(11,03+m.iii).Value  = dimstat(9,m.iii)
   .Cells(12,03+m.iii).Value  = dimstat(10,m.iii)
   .Cells(13,03+m.iii).Value  = dimstat(11,m.iii)
   .Cells(14,03+m.iii).Value  = dimstat(12,m.iii)
   .Cells(15,03+m.iii).Value  = dimstat(13,m.iii)
   .Cells(16,03+m.iii).Value  = dimstat(14,m.iii)
   
   .Cells(17,03+m.iii).Value  = dimstat(1,m.iii)+dimstat(2,m.iii)+dimstat(3,m.iii)+dimstat(4,m.iii)+dimstat(5,m.iii)+dimstat(6,m.iii)+;
    dimstat(7,m.iii)+dimstat(8,m.iii)+dimstat(9,m.iii)+dimstat(10,m.iii)+dimstat(11,m.iii)+dimstat(12,m.iii)+dimstat(13,m.iii)+;
    dimstat(14,m.iii)
    
   .Cells(18,03+m.iii).Value  = dimstat(1,m.iii)+dimstat(2,m.iii)+dimstat(3,m.iii)+dimstat(4,m.iii)+dimstat(5,m.iii)+dimstat(6,m.iii)+;
    dimstat(7,m.iii)+dimstat(8,m.iii)+dimstat(9,m.iii)+dimstat(10,m.iii)+dimstat(11,m.iii)+dimstat(12,m.iii)+dimstat(13,m.iii)+;
    dimstat(14,m.iii)+dimstat(15,m.iii)

  ENDFOR 

 ENDWITH 

 oDoc.SaveAs(DocName)
 oExcel.Visible = .t.

RETURN 

FUNCTION OneStat(para1)
 LOCAL m.nmon
 m.nmon = para1
 m.lperiod = STR(m.tyear,4)+PADL(m.nmon,2,'0')
 m.lpath = m.pbase+'\'+m.lperiod
 IF !fso.FolderExists(m.lpath)
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.lpath+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.lpath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(m.lpath+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.lpath+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.lpath+'\'+m.mcod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisom 
   LOOP 
  ENDIF 
  SELECT people
  SCAN 
   m.c_err = c_err
   DO CASE 
    CASE m.c_err = 'F0'
     dimstat(1,m.nmon) = dimstat(1,m.nmon) + 1
    CASE m.c_err = 'F1'
     dimstat(2,m.nmon) = dimstat(2,m.nmon) + 1
    CASE m.c_err = 'F2'
     dimstat(3,m.nmon) = dimstat(3,m.nmon) + 1
    CASE m.c_err = 'F3'
     dimstat(4,m.nmon) = dimstat(4,m.nmon) + 1
    CASE m.c_err = 'F4'
     dimstat(5,m.nmon) = dimstat(5,m.nmon) + 1
    CASE m.c_err = 'F5'
     dimstat(6,m.nmon) = dimstat(6,m.nmon) + 1
    CASE m.c_err = 'F6'
     dimstat(7,m.nmon) = dimstat(7,m.nmon) + 1
    CASE m.c_err = 'F7'
     dimstat(8,m.nmon) = dimstat(8,m.nmon) + 1
    CASE m.c_err = 'F8'
     dimstat(9,m.nmon) = dimstat(9,m.nmon) + 1
    CASE m.c_err = 'F9'
     dimstat(10,m.nmon) = dimstat(10,m.nmon) + 1
    CASE m.c_err = 'FF'
     dimstat(11,m.nmon) = dimstat(11,m.nmon) + 1
    CASE m.c_err = 'FP'
     dimstat(12,m.nmon) = dimstat(12,m.nmon) + 1
    CASE m.c_err = 'FL'
     dimstat(13,m.nmon) = dimstat(13,m.nmon) + 1
    CASE m.c_err = 'FD'
     dimstat(14,m.nmon) = dimstat(14,m.nmon) + 1
    OTHERWISE 
     dimstat(15,m.nmon) = dimstat(15,m.nmon) + 1
   ENDCASE 
  ENDSCAN 
  USE IN people 
  SELECT aisoms
 ENDSCAN 
 USE IN aisoms
RETURN 