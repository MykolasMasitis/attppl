PROCEDURE MakeActsLpu

IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре ятнплхпнбюрэ'+CHR(13)+CHR(10)+'тюикш ньханй дкъ кос?'+CHR(13)+CHR(10),4+32,'')=7
 RETURN 
ENDIF 

IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar") > 0
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 RETURN 
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "mcod") > 0
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 RETURN
ENDIF 

DotNameA = pTempl + '\svqqmmyy.dot'
DocNameA = pbase+'\'+m.gcperiod + '\sv' + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)

TRY 
 oWord = GETOBJECT(,"Word.Application")
CATCH 
 oWord = CREATEOBJECT("Word.Application")
ENDTRY 

oDocA = oWord.Documents.Add(dotnamea)

OldEscStatus = SET("Escape")
SET ESCAPE OFF 
CLEAR TYPEAHEAD 
 
GO TOP 

nLpu = 0
StartOfProc = SECONDS()
SCAN 
 m.mcod = mcod
 WAIT m.mcod WINDOW NOWAIT 

 IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+mcod)
  LOOP 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\people.dbf')
  LOOP 
 ENDIF 
  
 DocName = pbase+'\'+m.gcperiod+'\'+m.mcod+'\Pr'+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)

 IF fso.FileExists(DocName+'.pdf')
  fso.DeleteFile(DocName+'.pdf')
 ENDIF 

 =ActMoSmo(pbase+'\'+m.gcperiod+'\'+m.mcod, m.mcod, .f., .f.)

 IF fso.FileExists(DocName+'.doc')
* WAIT m.mcod WINDOW NOWAIT 
*  oDocA.Select
  oword.Selection.InsertFile(DocName+'.doc','',.f.,.f.,.f.)
  oword.Selection.InsertBreak(7)
 ENDIF 

 IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('бш унрхре опепбюрэ напюанрйс?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 

 WAIT CLEAR 
 nLpu = nLpu + 1
ENDSCAN 

USE 
USE IN sprlpu
 
EndOfProc = SECONDS()
LastOfProc = EndOfProc - StartOfProc
MeanTime = LastOfProc/nLpu

WAIT CLEAR 

SET ESCAPE &OldEscStatus
 
oDocA.SaveAs(DocNameA, 0)
oDocA.Close(0)
oWord.Quit

MESSAGEBOX(CHR(13)+CHR(10)+"напюанрйю гюйнмвемю!"+CHR(13)+CHR(10)+;
 "бяецн напюанрюмн кос   : "+TRANSFORM(nLpu, '9999999')+CHR(13)+CHR(10)+;
 "наыее бпелъ напюанрйх  : "+TRANSFORM(LastOfProc,'999.999')+" ЯЕЙ."+CHR(13)+CHR(10)+;
 "япедмее бпелъ напюанрйх: "+TRANSFORM(MeanTime,'999.999')+" ЯЕЙ."+CHR(13)+CHR(10),0+64,"")


RETURN 