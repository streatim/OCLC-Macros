'MacroName:949volume
'MacroDescription:949 volume
sub main
 Dim CS as object
 set CS = CreateObject("Connex.Client")  
   Dim locs(4) as string 
   Dim barcode(4) as string
   Dim tycod(4) as string
   Dim volinfo(4) as string
   Dim icode2
   Dim fl2$
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

   x=0
   tycod(1)="0"

CS.GetField "050", 1, lc_call
if lc_call="" then goto skipcall
callloc=""
if asc(mid$(lc_call,6,1))<=75 then callloc="c3rd"
if asc(mid$(lc_call,6,1))>=76 then callloc="c4th"

locs(1)=callloc
skipcall:
    
getinfo:
   x=x+1   

   locs(x)=InputBox$(culi$ + chr(10) + "#" & x & chr(10)&"Enter location:"&chr(10)&chr(10)&"Or enter 'done' to create 949.","949 volume",locs(1))
bci:  
   if locs(x)="" then goto out
   if locs(x)="done" or locs(x)="DONE" then x=x-1: goto prdline
   barcode(x)=InputBox$(culi$ + chr(10) + "#" & x & "  loc=" & locs(x) & chr(10)&chr(10)&"Scan barcode:","949 volume")
   if barcode(x)="" then goto out
   if len(barcode(x))<>14 then
   goto bci:
   end if
   tycod(x)=InputBox$(culi$ + chr(10) + "#" & x & "  loc=" & locs(x) & chr(10)&chr(10)&"Enter TY code:","949 volume",tycod(1))
   if tycod(x)="" then goto out
   volinfo(x)=InputBox$(culi$ + chr(10) + "#" & x & "  loc=" & locs(x) & "  ty=" & tycod(x) &chr(10)& chr(10)&"v=","949 volume")
   if volinfo(x)="" then goto out
   if x=4 then goto prdline
   culi$=culi$+"#" + x + "  loc="+locs(x)+ "  ty=" + tycod(x) + "  v=" + volinfo(x) + chr(10)
   goto getinfo

prdline: 

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

fl2$="949  *recs=b;bn=" & locs(1) & ";ins=" & userinit
for y= 1 to x
fl2$=fl2$ + ";i=" & barcode(y) & "/sta=" & istat & "/loc=" & locs(y) & "/ty=" & tycod(y) & "/v=" & volinfo(y)
if y=1 then fl2$=fl2$+"/i2=" & icode2 & "/b2=" & formatcode
next y
fl2$=fl2$+";"          

CS.addfieldline 1,fl2$
'CS.Cursorrow=1
'CS.Cursorcolumn=32
'cs.insertmode=false
out:
end sub
