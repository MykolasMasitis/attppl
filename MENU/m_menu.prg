PROCEDURE m_menu
SET SYSMENU TO

DEFINE PAD mmenu_1 OF _MSYSMENU PROMPT '\<���������� �� ���' COLOR SCHEME 3 KEY ALT+A, ""
DEFINE PAD mmenu_2 OF _MSYSMENU PROMPT '\<����������' COLOR SCHEME 3 KEY ALT+A, ""
DEFINE PAD mmenu_5 OF _MSYSMENU PROMPT '\<������' COLOR SCHEME 3 KEY ALT+F , ""
ON PAD mmenu_1 OF _MSYSMENU ACTIVATE POPUP popInfFrLpu
ON PAD mmenu_2 OF _MSYSMENU ACTIVATE POPUP popExp
ON PAD mmenu_5 OF _MSYSMENU ACTIVATE POPUP popTuneUp

DEFINE POPUP popInfFrLpu MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popInfFrLpu PROMPT '����� ��������� �� ��� (��� ���)'
DEFINE BAR 02 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 03 OF popInfFrLpu PROMPT '��������� �������������� �������'
DEFINE BAR 04 OF popInfFrLpu PROMPT '��������� �������������� �������'
DEFINE BAR 05 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 06 OF popInfFrLpu PROMPT '������� ������� �������'
DEFINE BAR 07 OF popInfFrLpu PROMPT '������������ ����� ������ ��� ���' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 08 OF popInfFrLpu PROMPT '������������ ���� ������ � ���' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 09 OF popInfFrLpu PROMPT '��������� ����� � ���' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 10 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 11 OF popInfFrLpu PROMPT '������������ ����� � ������' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 12 OF popInfFrLpu PROMPT '������������ ����� � ������ (����� ������)'
DEFINE BAR 13 OF popInfFrLpu PROMPT '������������ ������� ��� � ������' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 14 OF popInfFrLpu PROMPT '��������� ����� � ������' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 15 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 16 OF popInfFrLpu PROMPT '��������� ������ ������'
DEFINE BAR 17 OF popInfFrLpu PROMPT '������������ ������ ��� ��� (2-�� ����)' ;
 SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\'+'er'+m.qcod+PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),2)+'.dbf')
DEFINE BAR 18 OF popInfFrLpu PROMPT '������������ ����� ������ ��� �������� � ���' SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
DEFINE BAR 19 OF popInfFrLpu PROMPT '������� ��������� �������' ;
 SKIP FOR !fso.FileExists(pbase+'\'+gcperiod+'\'+'er'+m.qcod+PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),2)+'.dbf')
DEFINE BAR 20 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 21 OF popInfFrLpu PROMPT '�������� ������'
DEFINE BAR 22 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 23 OF popInfFrLpu PROMPT '�����'

ON SELECTION BAR 01 OF popInfFrLpu DO FORM MailView
ON SELECTION BAR 03 OF popInfFrLpu DO FORM MailTrash
ON SELECTION BAR 04 OF popInfFrLpu DO FORM MailDouble
ON SELECTION BAR 06 OF popInfFrLpu DO MakeSvPeople
ON SELECTION BAR 07 OF popInfFrLpu DO MakeETRLs
ON SELECTION BAR 08 OF popInfFrLpu DO MakeActsLpu
ON SELECTION BAR 09 OF popInfFrLpu DO SndToLpu
ON SELECTION BAR 11 OF popInfFrLpu DO MakeFOMS
ON SELECTION BAR 12 OF popInfFrLpu DO MakeFOMS2
ON SELECTION BAR 13 OF popInfFrLpu DO MakeActFOMS
ON SELECTION BAR 14 OF popInfFrLpu DO SendToFoms
ON SELECTION BAR 16	OF popInfFrLpu DO LoadErrFomsN
ON SELECTION BAR 17 OF popInfFrLpu DO MakeETRLs22
ON SELECTION BAR 18 OF popInfFrLpu DO MakeETRLsN
ON SELECTION BAR 19 OF popInfFrLpu DO MakeEtPeople
ON BAR 21 OF popInfFrLpu ACTIVATE POPUP popPrn
ON SELECTION BAR 23 OF popInfFrLpu CLEAR EVENTS 

DEFINE POPUP popPrn MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popPrn PROMPT '������ ���� ������'
ON SELECTION BAR 1 OF popPrn DO PackPrnSv

DEFINE POPUP popExp MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popExp PROMPT '���������� �������� �������'
DEFINE BAR 02 OF popExp PROMPT '������������ ������ �� ���� ��'
DEFINE BAR 03 OF popExp PROMPT '��������� ������ ���� ��'
DEFINE BAR 04 OF popExp PROMPT '������������� ������������'
ON SELECTION BAR 01 OF popExp DO FORM ExpView
ON SELECTION BAR 02 OF popExp DO PlWait01
ON SELECTION BAR 03 OF popExp DO PlWait02
ON BAR 04 OF popExp ACTIVATE POPUP AnnPpl

DEFINE POPUP AnnPpl MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 01 OF AnnPpl PROMPT '�������� ������ ��� ��'
DEFINE BAR 02 OF AnnPpl PROMPT '�������� �� � ������'
ON SELECTION BAR 01 OF AnnPpl do ChkAnnIpS
ON SELECTION BAR 02 OF AnnPpl do SndAnnIpS


DEFINE POPUP popTuneUp MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 01 OF popTuneUp PROMPT '����� ��������� �������' 
DEFINE BAR 02 OF popTuneUp PROMPT '��������� ������� ����������'
DEFINE BAR 03 OF popTuneUp PROMPT '��������� ������' 
DEFINE BAR 04 OF popTuneUp PROMPT '\-'
DEFINE BAR 05 OF popTuneUp PROMPT '�������������� �� ���'
DEFINE BAR 06 OF popTuneUp PROMPT '�������������� ������� ���'
DEFINE BAR 07 OF popTuneUp PROMPT '\-'
DEFINE BAR 08 OF popTuneUp PROMPT '��������� ����� ������'
DEFINE BAR 09 OF popTuneUp PROMPT '��������� ������� ����'
DEFINE BAR 10 OF popTuneUp PROMPT '\-'
DEFINE BAR 11 OF popTuneUp PROMPT '�������� ����� ������'
DEFINE BAR 12 OF popTuneUp PROMPT '������� CTRL-��'
DEFINE BAR 13 OF popTuneUp PROMPT '������� ����� ��������'
DEFINE BAR 14 OF popTuneUp PROMPT '������� BAK-�����'
DEFINE BAR 15 OF popTuneUp PROMPT '������� ALLPEOPLE'
DEFINE BAR 16 OF popTuneUp PROMPT '\-'
DEFINE BAR 17 OF popTuneUp PROMPT '������������� ��������� ������� ��'
DEFINE BAR 18 OF popTuneUp PROMPT '\-'
DEFINE BAR 19 OF popTuneUp PROMPT '��������� ������ � ������������'
DEFINE BAR 20 OF popTuneUp PROMPT '���������� �� �������'

ON SELECTION BAR 01 OF popTuneUp DO FORM SetPeriod
ON SELECTION BAR 02 OF popTuneUp DO FORM TuneBase
ON SELECTION BAR 03 OF popTuneUp goApp.doForm('set_print','settings')
ON SELECTION BAR 05 OF popTuneUp DO ComReind
ON SELECTION BAR 06 OF popTuneUp DO BasReind
ON SELECTION BAR 08 OF popTuneUp DO PackBD
ON SELECTION BAR 09 OF popTuneUp DO PackBDSv
ON SELECTION BAR 11 OF popTuneUp DO ZapEFiles
ON SELECTION BAR 12 OF popTuneUp DO DelAllCtrl
ON SELECTION BAR 13 OF popTuneUp DO DelAllZapros
ON SELECTION BAR 14 OF popTuneUp DO DelBakFiles
ON SELECTION BAR 15 OF popTuneUp DO DelAllPeople
ON SELECTION BAR 17 OF popTuneUp DO CorStruct
ON SELECTION BAR 19 OF popTuneUp DO AppMgfData
ON SELECTION BAR 20 OF popTuneUp DO ErrStat

SET SYSMENU AUTOMATIC
SET SYSMENU ON