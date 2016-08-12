FUNCTION OneReind(lcPath)

tc_mcod = SUBSTR(lcPath, RAT('\', lcPath)+1)

IF OpenFile("&lcPath\People", "People", "excl") == 0
 SELECT People 
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON n_pol TAG n_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX ON dr TAG dr
 INDEX ON c_err TAG c_err
 SET FULLPATH OFF 
 USE IN people 
ENDIF

IF fso.FileExists(lcPath+'\AllPeople.dbf')
IF OpenFile("&lcPath\AllPeople", "People", "excl") == 0
 SELECT People 
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON n_pol TAG n_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX ON dr TAG dr
 INDEX ON c_err TAG c_err
 SET FULLPATH OFF 
 USE IN people 
ENDIF
ENDIF 

RETURN 
