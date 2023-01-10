'MacroName:949
'MacroDescription:949


sub main
  Dim CS as object
  Set CS = CreateObject("Connex.Client")  
   Dim locs
   Dim barcode
   Dim icode2
   Dim tycod
   Dim formatcode as string
   Dim allformats as string
   Dim sData as String
   dim sdata2 as string
   dim m7a as string

 userinit = "bjt"
 icode2="-"  
 allformats="a    book"&chr$(10)&"0    periodical"&chr$(10)&"@    ebook"&chr$(10)&"g    DVD"&chr$(10)&"2    videotape"&chr$(10)&"3    laserdisc"&chr$(10)&"j    CD, music"&chr$(10)&"I    CD, spoken"&chr$(10)&"8    tape, music"&chr$(10)&"6    tape, spoken"&chr$(10)&"m    software"&chr$(10)&"9    website"&chr$(10)&"e    map"&chr$(10)&"c    music score"&chr$(10)&"t    thesis"&chr$(10)&"o    kit"&chr$(10)&"r    3D object"&chr$(10)&"4    slide"&chr$(10)&"h    record album"&chr$(10)

 CS.GetField "007", 1, sData
 sdata2=chr$(223)+"a "+mid$(sdata,6)+" "
 for x=2 to (len(sdata2)/5)+1
 if mid$(getfield (sdata2,x,chr$(223)),1,1)="a" then m7a=mid$(getfield (sdata2,x,chr$(223)),3,1)
 if mid$(getfield (sdata2,x,chr$(223)),1,1)="b" then m7b=mid$(getfield (sdata2,x,chr$(223)),3,1)
 if mid$(getfield (sdata2,x,chr$(223)),1,1)="e" then m7e=mid$(getfield (sdata2,x,chr$(223)),3,1)
 if mid$(getfield (sdata2,x,chr$(223)),1,1)="g" then m7g=mid$(getfield (sdata2,x,chr$(223)),3,1)
 next x
 CS.GetFixedField "Type", m6type
 CS.GetFixedField "blvl", m6blvl
 
 fortype="Unknown, please select"
 formatcode="-"
 if m6type="a" and m6blvl="m" then fortype="Book":formatcode="a"
 if m6type="a" and m6blvl="s" then fortype="Serial" & chr$(10) & chr$(10) & "Please choose whether item is a book (a) or periodical (0)":formatcode="a or 0"
 if m6type="m" and m6blvl="m" then fortype="Software":formatcode="m"
 if m6type="e" then fortype="Map":formatcode="e"
 if m6type="c" then fortype="Music Score":formatcode="c"
 if m6type="t" then fortype="Thesis":formatcode="t"
 if m6type="o" then fortype="Kit":formatcode="o"
 if m6type="r" then fortype="3D Object":formatcode="r"
 if m6type="g" and m7a="v" and m7e="v" then fortype="DVD":formatcode="g"
 if m6type="g" and m7a="v" and m7e="g" then fortype="Laserdisc":formatcode="3"
 if m6type="g" and m7a="v" and m7b="f" then fortype="Videocassette":formatcode="2"
 if m6type="g" and m7a="g" and m7e="s" then fortype="Slide":formatcode="4"
 if m6type="a" and m6blvl="m" and m7a="c" and m7b="r" then fortype="ebook or website":formatcode="@ or 9"
 if m6type="m" and m6blvl="m" and m7a="c" and m7b="r" then fortype="ebook or website":formatcode="@ or 9" 
 if m6type="j" and m7a="s" and m7b="d" and m7g="g" then fortype="CD, Music":formatcode="j"
 if m6type="i" and m7a="s" and m7b="d" and m7g="g" then fortype="CD, Spoken":formatcode="i"
 if m6type="j" and m7a="s" and m7b="s" then fortype="tape, Music":formatcode="8"
 if m6type="i" and m7a="s" and m7b="s" then fortype="tape, spoken":formatcode="6"
 if m6type="i" and m7a="s" and m7b="d" and m7g<>"g" then fortype="Record Album":formatcode="h"
 if m6type="j" and m7a="s" and m7b="d" and m7g<>"g" then fortype="Record Album":formatcode="h"

CS.GetField "050", 1, lc_call
if lc_call="" then goto skipcall
callloc=""
if asc(mid$(lc_call,6,1))<=75 then callloc="c3rd"
if asc(mid$(lc_call,6,1))>=76 then callloc="c4th"
skipcall:

   locs=InputBox$("Enter location:","949",callloc)
   if locs="" then goto out
bci:  barcode=InputBox$("Scan barcode:","949")
   if barcode="" then goto out
   if len(barcode)<>14 then goto bci:
   tycod=InputBox$("Enter TY code:","949","0")
   if tycod="" then goto out   
noformat:
   curformat=formatcode
   formatcode=inputbox$("Format:"&chr$(10)&chr$(10)&"Suggested code: "& fortype & chr$(10) & chr$(10) & "Enter 'list' to see available codes.","949",formatcode)
   if formatcode="" then goto out
   if formatcode="list" then msgbox (allformats): formatcode=curformat: goto noformat:
   if len(formatcode) <> 1 then goto noformat:

noicode2:
   icode2=inputbox$("Icode2:" & chr$(10) & chr$(10) & "Choose from:" & chr$(10) & "-    None" & chr$(10) & "a    CONTENT ADDED" & chr$(10) & "b    SUBJECT ADDED" & chr$(10)& "c    NOTE/SUB ADDED","949",icode2)
   if icode2="" then goto out
   if len(icode2) <> 1 then goto noicode2
   
i_status:
   istat=inputbox$("Item Status:" & chr$(10) & chr$(10) & "Choose from:" & chr$(10) & "-   Available" & chr$(10) & "p   In Process","949","p")
   if istat="" then goto i_status
   if istat="p" then goto make949
   if istat="-" then goto make949
   goto i_status

make949:      
CS.addfieldline 1,"949  *recs=b;bn=" & locs & ";ins=" & userinit & ";i=" & barcode & "/sta=" & istat & "/loc=" & locs & "/ty=" & tycod & "/i2=" & icode2 & "/b2=" & formatcode & ";"
'CS.Cursorrow=1
'CS.Cursorcolumn=32

out:
end sub
