Services.StartTransaction "tr1"
 
mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount
For i = 1 To mrowcount Step 1
Datatable.SetCurrentRow(i)
Modexe=Datatable("ModuleExe","Action1")
'msgbox Modexe
If Modexe="Y" Then
        Modid=Datatable("ModuleID","Action1")
        msgbox Modid
        trowcount=datatable.GetSheet("Action2").GetRowCount
        msgbox trowcount
        For j=1 To trowcount Step 1
    Datatable.SetCurrentRow(j)
    If Modid=Datatable("ModuleID","Action2") and Datatable("TestCaseExecution","Action2")="Y" then
    testcaseid=Datatable("TCaseId","Action2")
    msgbox testcaseid
        tsrowcount=Datatable.GetSheet("Action3").GetRowCount
        msgbox tsrowcount
        For k = 1 to tsrowcount Step 1
        datatable.SetCurrentRow(k)
        If testcaseid=Datatable("TestCaseId","Action3") Then
        keyword=Datatable("Keyword","Action3")
        msgbox keyword
        Select case (keyword)

        Case "ln"
        Call Login("john","hp")

        Case "ca"
        Call CloseApp()

        Case "oo"
        Call OpenOrder()
        Case "uo"
        Call UpdateOrder()
        Case "lnd"
        
        drowcount=datatable.GetSheet("Action4").GetRowCount
        For l = 1 To drowcount Step 1
        	Datatable.SetCurrentRow(l)
        	Call Login(Datatable("Username","Action4"),Datatable("Password","Action4"))
        	Call CloseApp()
        Next
        Case "ood"
        orowcount=datatable.GetSheet("Action4").GetRowCount
        For m = 1 To orowcount Step 1
        	Datatable.SetCurrentRow(m)
        	Call OpenOrder(Datatable("OrderNo","Action4"))
        Next

        End  Select

        End If

        Next


    End If

    Next




End If
 




Next


 

Services.EndTransaction "tr1"
 @@ hightlight id_;_2430074_;_script infofile_;_ZIP::ssf27.xml_;_
