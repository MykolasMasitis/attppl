PROCEDURE mgfrecalc
 IF MESSAGEBOX(CHR(13)+CHR(10)+'оепеявхрюрэ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
* IF !fso.FolderExists(m.pbase+'\'+m.gcperiod)
*  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер дхпейрнпхъ оепхндю!'+CHR(13)+CHR(10),0+16,'')
*  RETURN 
* ENDIF 
 
* IF !fso.FileExists(m.pbase+'\'+m.gcperiod+'\aisoms.dbf')
*  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик AISOMS.DBF!'+CHR(13)+CHR(10),0+16,'')
*  RETURN 
* ENDIF 

 orec = RECNO()
 
 SCAN 
  m.ad_mgf = m1824 + f1824 +m2534+f2534+m3544+f3544+m4559+f4559+m6068+f5564+m69+f65
  m.ch_mgf = ch01m + ch01f +ch14m+ch14f+ch514m+ch514f+ch1517m+ch1517f
  
  IF ad_mgf != m.ad_mgf
   REPLACE ad_mgf WITH m.ad_mgf
  ENDIF 
  IF ch_mgf != m.ch_mgf
   REPLACE ch_mgf WITH m.ch_mgf
  ENDIF 
 ENDSCAN 
 
 GO (orec)
 
 
RETURN 