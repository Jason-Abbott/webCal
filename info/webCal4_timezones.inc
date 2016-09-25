<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/26/2000

dim arZones			' array of time zones

Function listzones()
	dim x			' loop counter
	dim strHTML		' build select list
	
	for x = 0 to UBound(arZones)
		strHTML = strHTML & "<option value=" & arZones(x)(1) _
			& ">" & arZones(x)(0) & vbCrLF
	next
	listzones = strHTML
End Function

' WINTER

arZones = Array( _
	Array("(GMT-12) Eniwetok, Kwajalein",-12), _
	Array("(GMT-11) Midway Island, Samoa",-11), _
	Array("(GMT-10) Hawaii",-10), _
	Array("(GMT-9) Alaska",-9), _
	Array("(GMT-8) Pacific Time, Tijuana",-8), _
	Array("(GMT-7) Mountain Time",-7), _
	Array("(GMT-6) Central Time",-6), _
	Array("(GMT-5) Eastern Time",-5), _
	Array("(GMT-4) Atlantic Time",-4), _
	Array("(GMT-3)",-3), _
	Array("(GMT-2)[AT] Azores",-2), _
	Array("(GMT-1)[WAT] West Africa",-1), _
	Array("(GMT )[GMT] England, Scotland",0), _
	Array("(GMT+1)",1), _
	Array("(GMT+2)",2), _
	Array("(GMT+3)",3), _
	Array("(GMT+4)",4), _
	Array("(GMT+5)",5), _
	Array("(GMT+6)",6), _
	Array("(GMT+7)",7), _
	Array("(GMT+8)",8), _
	Array("(GMT+9)",9), _
	Array("(GMT+10)",10), _
	Array("(GMT+11)",11), _
	Array("(GMT+12)",12) _
)

-3:00
 c
 Charlie
   Brasilia, Brazil
Buenos Aires, Argentina
Georgetown, Guyana 
-4:00
 d
   Delta
 AST - Atlantic Standard Caracas
La Paz 
-5:00
 e
   Echo
 EST - Eastern Standard. Bogota
Lima, Peru
New York, NY, USA 
-6:00
 f
   Foxtrot
 CST - Central Standard Mexico City, Mexico
Saskatchewan, Canada 
-7:00
 g
   Golf
 MST - Mountain Standard.   
-8:00
 h
   Hotel
 PST - Pacific Standard Los Angeles, CA, USA 
-9:00
 j
 Juliet 
 YST - Yukon Standard.   
-10:00
 k
   Kilo
 AHST - Alaska-Hawaii Standard
CAT - Central Alaska
HST - Hawaii Standard 
EAST - East Australian Standard
   
-11:00
 l
   Lima
 NT - Nome.   
-12:00
 m
   Mike
 IDLW - International Date Line West   
+1:00
 n
   November
 CET - Central European
FWT - French Winter
MET - Middle European
MEWT - Middle European Winter
SWT - Swedish Winter. Paris, France
Berlin, Germany
Amsterdam, The Netherlands
Brussels, Belgium
Vienna, Austria
Madrid, Spain
Rome, Italy
Bern, Switzerland
Stockholm, Sweden
Oslo, Norway 
+2:00
 o
   Oscar
 EET - Eastern European, Russia Zone 1 Athens, Greece
Helsinki, Finland
Istanbul, Turkey
Jerusalem, Israel
Harare, Zimbabwe 
+3:00
 p
   Papa
 BT - Baghdad, Russia Zone 2 Kuwait
Nairobi, Kenya
Riyadh, Saudi Arabia
Moscow, Russia
Tehran, Iran 
+4:00
 q
   Quebec
 ZP4 - Russia Zone 3 Abu Dhabi, UAE
Muscat
Tblisi
Volgograd
Kabul 
+5:00
 r
   Romeo
 ZP5 - Chesapeake Bay   
+6:00
 s
   Sierra
 ZP6 - Chesapeake Bay   
+7:00
 t
   Tango
 WAST - West Australian Standard.   
+8:00
 u
   Uniform
 CCT - China Coast, Russia Zone 7   
+9:00
 v
   Victor
 JST - Japan Standard, Russia Zone 8   
+10:00
 w
   Whiskey
 GST - Guam Standard, Russia Zone 9   
+11:00
 x
   X-ray
     
+12:00
 y
   Yankee
 IDLE - International Date Line East
NZST - New Zealand Standard
NZT - New Zealand Wellington, New Zealand
Fiji
Marshall Islands 


SUMMER/DAYLIGHT SAVINGS
There are many Daylight Savings Time (DST) time-zones, indexed by the zone hour value relative to both Greenwich Mean Time (GMT) and British Summer Time (BST).  

-1:00
 +0:00
  
 BST - British Summer. 
+0:00
 +1:00
  
   
+1:00
 +2:00
  
   
+2:00
 +3:00
  
   
+3:00
 +4:00
  
 ADT - Atlantic Daylight 
+4:00
 +5:00
  
 EDT - Eastern Daylight. 
+5:00
 +6:00
  
 CDT - Central Daylight 
+6:00
 +7:00
  
 MDT - Mountain Daylight. 
+7:00
 +8:00
  
 PDT - Pacific Daylight. 
+8:00
 +9:00
  
 YDT - Yukon Daylight. 
+9:00
 +10:00
  
 HDT - Hawaii Daylight. 
+10:00
 +11:00
  
   
+11:00
 +12:00
  
   
-2:00
 -1:00
  
 MEST - Middle European Summer
MESZ - Middle European Summer
SST - Swedish Summer
FST - French Summer. 
-3:00
 -2:00
  
   
-4:00
 -3:00
  
   
-5:00
 -4:00
  
   
-6:00
 -5:00
  
   
-7:00
 -6:00
  
   
-8:00
 -7:00
  
 WADT - West Australian Daylight 
-9:00
 -8:00
  
   
-10:00
 -9:00
  
   
-11:00
 -10:00
  
 EADT - Eastern Australian Daylight 
-12:00
 -11:00
  
   
-13:00
 -12:00
  
 NZDT - New Zealand Daylight. 




 