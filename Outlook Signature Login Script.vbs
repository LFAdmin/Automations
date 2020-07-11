'Script is written by Prabhat Nigam
'For OAG only
'Copy Write Prabhat Nigam
'Version 3.0 - updated empty variable and user name
On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strFName = objUser.givenName
strI = objuser.initials
strLName = objuser.sn
strTitle = objUser.Title
strCompany = objUser.Company
strStreet = objuser.streetaddress
strCity = objuser.l
strstate = objuser.st
strzip = objuser.postalCode
Stremail = objuser.mail
strPhone = objUser.telephoneNumber
StrMobile = objuser.mobile
StrFax = objuser.facsimileTelephoneNumber
Strweb = wWWHomePage


Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Arial"
objSelection.Font.Size = 11
objSelection.Font.Bold = True
objSelection.TypeText strFName
If strI <> "" then
  objSelection.TypeText  "  " & strI
End If
objSelection.TypeText "  " & strLName
objSelection.TypeText Chr(11)
objSelection.TypeText strTitle
objSelection.TypeText Chr(11)
objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10.5
objSelection.Font.Bold = True
objSelection.Font.Italic = True
objSelection.Font.Color = vbBlue
objSelection.TypeText strCompany
objSelection.TypeText Chr(11)
objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10
objSelection.Font.Bold = False
objSelection.Font.Italic = False
objSelection.Font.Color = vbBlack
objSelection.TypeText strStreet
objSelection.TypeText Chr(11)
objSelection.TypeText strCity & ", " & strState & ", " & strZip
objSelection.TypeText Chr(11)
If strPhone <> "" then
  objSelection.TypeText strPhone & " Office"
  objSelection.TypeText Chr(11)
End If
If strFax <> "" then
  objSelection.TypeText strFax & " Fax"
  objSelection.TypeText Chr(11)
End If
If strMobile <> "" then
  objSelection.TypeText strMobile & " Mobile"
  objSelection.TypeText Chr(11)
End If
objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10
objSelection.Font.Color = vbBlue
objSelection.Font.underline = True
objSelection.TypeText stremail
objSelection.TypeText Chr(11)
objSelection.TypeText "http://www.ag.virginia.gov"
bjSelection.Hyperlinks.Add objShape, "http://www.ag.virginia.gov"
objSelection.TypeText Chr(11)
set objShape = objSelection.InlineShapes.AddPicture("\\oag.state.va.us\SysVol\oag.state.va.us\Policies\{E520082F-D07A-42BB-B4DB-C0C67A068446}\User\Scripts\Logon\OAGSealsSig\oag seal.jpg")
objSelection.Hyperlinks.Add objShape, "http://www.ag.virginia.gov"
objSelection.TypeText Chr(11)
set objShape = objSelection.InlineShapes.AddPicture("\\oag.state.va.us\SysVol\oag.state.va.us\Policies\{E520082F-D07A-42BB-B4DB-C0C67A068446}\User\Scripts\Logon\OAGSealsSig\Fb.jpg")
objSelection.Hyperlinks.Add objShape, "https://www.facebook.com/AGMarkHerring"
set objShape = objSelection.InlineShapes.AddPicture("\\oag.state.va.us\SysVol\oag.state.va.us\Policies\{E520082F-D07A-42BB-B4DB-C0C67A068446}\User\Scripts\Logon\OAGSealsSig\Twitter.jpg")
objSelection.Hyperlinks.Add objShape, "https://twitter.com/AGMarkHerring"

Set objSelection = objDoc.Range()

objSignatureEntries.Add "AD Signature", objSelection
objSignatureObject.NewMessageSignature = "AD Signature"
objSignatureObject.ReplyMessageSignature = "AD Signature"

objDoc.Saved = True
objWord.Quit
