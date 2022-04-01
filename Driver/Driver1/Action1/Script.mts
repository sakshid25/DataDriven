'Datatable.AddSheet "Module"
'Datatable.ImportSheet "‪C:\Users\sfjbs\Documents\UFT One\Organizer\Oraganizer.xlsx",1,"Module"
mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount

For  i = 1 To mrowcount Step 1
	
Datatable.SetCurrentRow(i)

Modexe=Datatable("ModuleExe","Action1")

msgbox Modexe
If modexe="Y" Then
	
	Modid=Datatable("ModuleID","Action1")
	
	msgbox Modid
	
	trowcount=datatable.GetSheet("Action2").GetRowCount
	
	msgbox trowcount
	
	For j=1  To trowcount Step 1
	Datatable.SetCurrentRow(j)
	If Modid=Datatable("ModuleID","Action2") and Datatable("TestCaseExe","Action2")="Y" Then
	testcaseid=Datatable("TestCaseID","Action2")
	msgbox testcaseid	
	
	
	tsrowcount=datatable.GetSheet("Action3").GetRowCount
	
	msgbox tsrowcount
	
	For k=1  To tsrowcount Step 1
	Datatable.SetCurrentRow(k)
	If testcaseid=Datatable("TestCaseID","Action3") Then
	keyword=Datatable("Keyword","Action3")
	msgbox keyword	
	
	 Select case (keyword)
 
        Case "ln"
        Call Login("john","hp")
 
        Case "ca"
        Call Closeapp()
 
        Case "oo"
        Call OpenOrder("5")
        
        Case "uo"
        Call UpdateOrder()
        
        Case "lnd"
        drowcount = datatable.GetSheet("Action4").GetRowCount
        For l=1 = 1 To drowcount Step 1
        	datatable.SetCurrentRow(l)
        	Call login(datatable("username","Action4"),datatable("password","Action4"))
        	Call Closeapp()
        Next
        
        Case "ood"
        orrowcount = datatable.GetSheet("Action4").GetRowCount
        For m=1 To drowcount Step 1
        	datatable.SetCurrentRow(m)
        	Call openorder(datatable("orderno","Action4"))
        	Call Closeapp()
       
        
        
        End  Select
 
        End If
 
        Next
 

    End If
 
    Next
 
 

End If
	
	
Next

 @@ hightlight id_;_1893831824_;_script infofile_;_ZIP::ssf2.xml_;_
