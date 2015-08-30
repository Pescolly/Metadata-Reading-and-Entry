Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = WScript.CreateObject("WScript.Shell")

sh.Run(Chr(34) & "C:\Documents and Settings\All Users\Start Menu\MVF-9.lnk" & Chr(34))
WScript.Sleep(1000)
sh.SendKeys "AKARAMIAN"
sh.SendKeys "{ENTER}"
WScript.Sleep(100)
sh.SendKeys "******"
sh.SendKeys "{ENTER}"
WScript.Sleep(100)
sh.SendKeys "D"
WScript.Sleep(100)
sh.SendKeys "E"
WScript.Sleep(100)
sh.SendKeys "4"
WScript.Sleep(100)
sh.SendKeys "6"
WScript.Sleep(100)
sh.SendKeys "{TAB}"

Set fileString = fso.OpenTextFile("X:\AssetManagement\Work\scriptTestFile.txt", 1)
'Do Until listFile.AtEndOfStream
sh.SendKeys fileString.ReadLine
sh.SendKeys "{TAB}{TAB}"
sh.SendKeys "M"				'enter as master
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		'enter production number
sh.SendKeys "{TAB}"		
sh.SendKeys fileString.ReadLine		'enter spec1
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		'enter spec2
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		'enter standard (Pal Ntsc 3-2398, 1-5994, 2-1080i50
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		'enter tape size "PRH"
sh.SendKeys "{TAB}"
sh.SendKeys "UNIJA"			'enter UNIJA as owner
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		'enter aspect ratio (4x3: 3=1.33,L=1.78, D=2.35,G=2.40, 16x9: H=1.33, 8=1.78,J=1.85,p=2.40)
sh.SendKeys "{TAB}"
sh.SendKeys "UDS"			'enter recording place
sh.SendKeys "{TAB}"
sh.SendKeys "Y"				'yes to notes/comments
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		'trt minutes
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		'trt seconds
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		'DA number
sh.SendKeys "{TAB}"
sh.SendKeys fileString.ReadLine		't-code D=30DF, N=30NDF, P=24, E=25
sh.SendKeys "{TAB}"

'Ok button to confirm and move onto next entry
'Loop