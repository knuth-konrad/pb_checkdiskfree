'------------------------------------------------------------------------------
'Purpose  : Check a drive's free space and return a %ERRORLEVEL%
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 15.07.2016
'   Source: -
'  Changed: 15.05.2017
'           - Application manifest added.
'------------------------------------------------------------------------------
#Compile Exe ".\CheckDiskFree.exe"
#Option Version5
#Dim All

#Break On
#Debug Error On
#Tools Off

DefLng A-Z

%VERSION_MAJOR = 1
%VERSION_MINOR = 0
%VERSION_REVISION = 2

' Version information resource
#Include ".\ChkDskFreeRes.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include "win32api.inc"
#Include "sautilcc.inc"

Declare Sub ShowHelp
Declare Function CalcVal (ByVal sValue As String) As Quad
'------------------------------------------------------------------------------
'*** Variabels ***
'------------------------------------------------------------------------------

Function PBMain()
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 15.07.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local i As Long            'Loop counter
   Local sCommand As String   'Command line
   Local sFile As String      'File to check
   Local sTemp As String
   Local qudVal As Quad, qudSize As Quad
   Local lCompare As Long     ' Comparison to perform


   ' Application intro
   ConHeadline "CheckDiskFree", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2012-2016", $COMPANY_NAME
   Con.StdOut ""

   '** Parse command line
   sCommand = Command$
   If Len(Trim$(sCommand)) < 1 Then
      ShowHelp
      Function = 100
      Exit Function
   End If

   For i = 1 To ArgC()
      sTemp = ArgV(i)
      If LCase$(Left$(sTemp, 2)) = "/d" Then
         sFile = Trim$(Mid$(sTemp, 4))
      Else
         Select Case LCase$(Left$(Trim$(sTemp), 2))
         Case "/s"
            qudVal = CalcVal(Mid$(sTemp, 4))
         Case "/c"
            lCompare = Val(Mid$(sTemp, 4))
         Case Else
            ShowHelp
            Function = 100
            Exit Function
         End Select
      End If
   Next i

   If lCompare >= -2 And lCompare <= 2 Then
      Try
         qudSize = DiskFree(sFile)
      Catch
         Function = 50
      End Try
      StdOut "Comparing: " & sFile & " ";
      Select Case lCompare
      Case 0
         StdOut "=";
         If qudSize = qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case 2
         StdOut ">";
         If qudSize > qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case 1
         StdOut ">=";
         If qudSize >= qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case -2
         StdOut "<";
         If qudSize < qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case -1
         StdOut "<=";
         If qudSize <= qudVal Then
            Function = 0
            sTemp = "Passed!"
         Else
            Function = 1
            sTemp = "Failed!"
         End If
      Case Else
         Function = 100
      End Select
      StdOut " " & Extract$(FormatNumber(qudVal), ",") & " bytes."
   Else
      Function = 100
   End If

   StdOut sTemp

End Function
'---------------------------------------------------------------------------

Sub ShowHelp

   StdOut "CheckDiskFree"
   StdOut "-------------"
   StdOut "CheckDiskFree determines if a disk's free space fits into the given comparison."
   StdOut ""
   StdOut "Usage:   CheckDiskFree /d=<disk/drive> /s=<size> /c=<compare argument>"
   StdOut "i.e.     CheckDiskFree /d=d: /s=1000000 /c=1"
   StdOut ""
   StdOut "Parameters"
   StdOut "----------"
   StdOut "/d   = Drive to check."
   StdOut "/s   = Size to check against. Valid fomats:"
   StdOut "       1000 - figures only = Bytes."
   StdOut "        2kb - Kilobytes where 1kb = 1024 Bytes."
   StdOut "        5mb - Megabytes where 1mb = 1024 Kilobytes."
   StdOut "        3gb - Gigabytes where 1gb = 1024 Megabytes."
   StdOut "        1tb - Terrabytes where 1tb = 1024 Gigabytes."
   StdOut "/c   = Comparison to perform. Valid operators:"
   StdOut "     - -2 free space must be lesser than given size."
   StdOut "     - -1 free space must be lesser than or equal to given size."
   StdOut "     -  0 free space must equal to given size."
   StdOut "     -  1 free space must be greater than or equal to given size."
   StdOut "     -  2 free space must be greater than given size."
   StdOut ""
   StdOut "CheckDiskFree returns the following DOS error levels upon exit, determing success"
   StdOut "or failure of the operation:"
   StdOut "  0  = Drive free space has passed comparison."
   StdOut "  1  = Drive free space size has *not* passed comparison."
   StdOut " 50  = Drive not found/doesn't exist"
   StdOut "100  = Invalid/missing command line parameter"
   StdOut "255  = Other (application) error"

End Sub
'---------------------------------------------------------------------------

Function CalcVal (ByVal sValue As String) As Quad
'------------------------------------------------------------------------------
'Purpose  : Calculate the multiplication factor from a unit's name
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 15.07.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   sValue = LCase$(sValue)
   Select Case Right$(sValue, 2)
   Case "kb"
      CalcVal = Val(sValue) * 1024&&
   Case "mb"
      CalcVal = Val(sValue) * 1024&&^2
   Case "gb"
      CalcVal = Val(sValue) * 1024&&^3
   Case "tb"
      CalcVal = Val(sValue) * 1024&&^5
   Case Else
      CalcVal = Val(sValue)
   End Select

End Function
'---------------------------------------------------------------------------
