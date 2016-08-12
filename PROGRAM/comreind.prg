PROCEDURE ComReind

WAIT "Ïåðåèíäåêàöèÿ COMMON..." WINDOW NOWAIT 
IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\UsrLpu', "usrlpu", "excl") == 0
 SELECT UsrLpu
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON mcod TAG mcod
 INDEX ON lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

m.lWasUsed = .F.
IF USED('Users')
 m.lWasUsed = .T.
 SELECT Users
 COUNT FOR ISRLOCKED() TO m.nLocked
 IF m.nLocked > 0
  SELECT name FROM Users WHERE !ISRLOCKED() INTO CURSOR wlck
  SELECT wlck
  INDEX on name TAG name 
  SET ORDER TO name 
 ENDIF 
 USE IN users
ENDIF 
IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
 SELECT Users 
 INDEX ON name TAG name 
 USE 
ENDIF 
IF m.lWasUsed=.T.
 =OpenFile(pCommon+'\Users', 'Users', 'shar', 'name')
 IF USED('wlck')
  SELECT Users
  SET RELATION TO name INTO wlck
  SCAN 
   IF EMPTY(wlck.name)
    RLOCK()
   ENDIF 
  ENDSCAN 
  SET RELATION OFF INTO wlck
  USE IN wlck
 ENDIF 
ENDIF 

IF !fso.FileExists(pCommon+'\Users.cdx')
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT Users 
  INDEX ON name TAG name 
  USE 
 ENDIF 
ENDIF 

WAIT "Ïåðåèíäåêñàöèÿ ëîêàëüíîé ÍÑÈ..." WINDOW NOWAIT 
IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx', "admokr", "excl") == 0
 SELECT admokr
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON cokr TAG cokr
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoerzxx', "OsoERZ", "excl") == 0
 SELECT OsoERZ
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON Ans_r TAG Ans_r
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

*IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoree', "osoree", "excl") == 0
* SELECT osoree
* SET FULLPATH OFF 
* WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON d_type TAG d_type
* SET FULLPATH OFF 
* USE
* WAIT CLEAR 
*ENDIF

*IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\smo', "smo", "excl") == 0
* SELECT smo
* SET FULLPATH OFF 
* WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON code TAG code
* SET FULLPATH OFF 
* USE
* WAIT CLEAR 
*ENDIF

IF OpenFile(pcommon+'\smo', "smo", "excl") == 0
 SELECT smo
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', "sprabo", "excl") == 0
 SELECT sprabo
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON abn_name TAG abn_name
 INDEX ON object_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "excl") == 0
 SELECT sprlpu
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 INDEX ON fil_id TAG fil_id
 INDEX ON mcod TAG mcod
 INDEX ON cokr TAG cokr
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\street', "street", "excl") == 0
 SELECT street
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON ul TAG ul
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tiplpu', "tiplu", "excl") == 0
 SELECT tiplu
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON lpu_id TAG lpu_id
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

IF !fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\errors.dbf')
 IF fso.FileExists(pcommon+'\errors.dbf')
  fso.CopyFile(pcommon+'\errors.dbf', pbase+'\'+gcperiod+'\'+'nsi'+'\errors.dbf')
 ENDIF 
ENDIF 

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\errors', "errors", "excl") == 0
 SELECT errors
 SET FULLPATH OFF 
 WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
 INDEX ON code TAG code
 SET FULLPATH OFF 
 USE
 WAIT CLEAR 
ENDIF

*IF OpenFile(pcommon+'\pilot', "pilot", "excl") == 0
* SELECT pilot
* SET FULLPATH OFF 
* WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON lpu_id TAG lpu_id
* INDEX ON mcod TAG mcod 
* SET FULLPATH OFF 
* USE
* WAIT CLEAR 
*ENDIF

*IF OpenFile(pcommon+'\tpn', "tpn", "excl") == 0
* SELECT tpn
* SET FULLPATH OFF 
* WAIT "ÈÍÄÅÊÑÈÐÎÂÀÍÈÅ ÔÀÉËÀ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
* INDEX ON mcod TAG mcod
* INDEX on lpuid TAG lpuid
* SET FULLPATH OFF 
* USE
* WAIT CLEAR 
*ENDIF
