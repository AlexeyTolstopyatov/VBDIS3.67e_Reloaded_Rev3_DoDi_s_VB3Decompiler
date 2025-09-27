### Note by me (CoffeeLake)

This is a last found version of DoDi's VB decompiler
which sometimes I'm trying to demangle and restrict.
My main task of it - determine Microsoft Visual Basic 2.0/3.0
specifics and runtime behavior.

The information about Visual Basic 4.0 exists here and in demangled
ActiveX objects of VB SemiDecompiler by VBGamer45. 

I've published my instance of found software under GNU GPL license
because this information already public. And my modifications of this
will be public too. 

This is a big legacy which I want to demangle and rebase to 
modern platforms. 
Despite the fact that I've been writing in Visual Basic 5 since childhood, 
I'm trying to get rid of it completely in the future and rebase code on the 
modern platforms. 

### Note by previous owner (CW2K@gmx.de)

DoDi is a Decompiler 3.67e for VB2 & VB3.
That's the lastest Version (1997) I could found.

DoDi itself is written in VB3 but has a 'protection'
to protect itself an other VB3 programm to reveal it's
sourcecode. So it can't be decompile it self.

Just for fun I 'analysed' this protection to remove it.
So after rebuilding what the protection had destroyed, 
decompiling and debugging of the decompilered source I got it 
running for it source own decompiled source.
...and also improved it so now it will also decompile
'Protected VB3-exe'.

So happy VB3-Decompiling !



So what VBGuard does:
1. Remove all the unnecessary uneven Control resource which look like this
```
00006500: 05 46 6F 72 6D 31 08 53 53 46 72 61 6D 65 31 08  .Form1.SSFrame1.
00006510: 63 6F 6E 74 72 6F 6C 32 0B 74 78 74 5F 46 69 6C  control2.txt_Fil
00006520: 65 4F 75 74 08 63 6F 6E 74 72 6F 6C 34 08 63 6F  eOut.control4.co
00006530: 6E 74 72 6F 6C 35 0B 74 78 74 5F 46 69 6C 65 52  ntrol5.txt_FileR
00006540: 65 73 08 63 6F 6E 74 72 6F 6C 37 08 63 6F 6E 74  es.contr
```

For each form `VB3` will add two Resource Items to the exe
(use a resviewer like exescope to view it)
RESID	Descibtion
4	Form1
5	Form1 'unnecessary data'
6	Form2
7	Form2 'unnecessary data'
8	Form3  ...
Of course these unnecessary data are great for the decompilation and make it more 'speakable'.


2. Remove of FormNameData in the `'MainFormStruct`
```
Before
00004C40: 44 49 41 4C 4F 47 2E 56 42 58 00 46 0A 04 80 54  DIALOG.VBX.F.. T
00004C50: 00 FF 01 14 B8 0A 46 30 30 30 30 2E 46 52 4D 00  .... .F0000.FRM.
00004C60: 00 00 46 0A 06 80 55 00 FF 01 10 80 0C 46 30 30  ..F.. U.... .F00
00004C70: 30 31 2E 46 52 4D 00 00 00 58 00 00 00 1A 00 1B  01.FRM...X......
00004C80: 00 D7 00 00 00 00 58 01 00 00 1C 00 1D 00 D7 00  . ....X....... .
```
And 
```
After
00004C40: 44 49 41 4C 4F 47 2E 56 42 58 00 46 0A 04 80 54  DIALOG.VBX.F.. T
00004C50: 00 FF 01 14 B8 00 00 00 00 00 00 00 00 00 00 00  .... ...........
```

3. Change the order of the segments. Place 3. segments at the beginning 
	so if old order was 1,2,3,4,5... new it will be 3,1,2,4,5...
   In the Decompiler is hardcoded to start very time with segments 3.
   If the order is change it will read garbage data and crash.
   
4. It does a case sensitive compare oft the string 
   'VBRUN300' to detect if it's a VB-exe. VbGuard changes
    this to 'vbrun300'. Windows will still find vbrun300.DLL
    but DoDi will say: 'No Visual Basic Exe'

<CW2K@gmx.de>

### Notes

1. Read 16bit-Token from Seg
	Token needs to be between `1..32511` (`0x7EFF`)
	VBDIS_ControlToken(Token)
	
	1..0x2A55(0x7EFF \ 3)
	VBDIS_FlagToken.FileData(intTokens \ 3)
	...


