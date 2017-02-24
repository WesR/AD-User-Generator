Set objRootDSE = GetObject("LDAP://rootDSE") 
 
Set objContainer = GetObject("LDAP://cn=Users," & _ 
    objRootDSE.Get("defaultNamingContext")) 
  
For i = 1 To 10
	Set objLeaf = objContainer.Create("User", "cn=UserNo" & i)
	objLeaf.Put "sAMAccountName", "HalCorpUser" & i
	objLeaf.SetInfo
Next 
  
WScript.Echo "Accounts Created." 
