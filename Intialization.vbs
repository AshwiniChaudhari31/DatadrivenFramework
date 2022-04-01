
Dim objuft

Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.Open("C:\Users\Lenovo\Documents\DataDrivenFramework\Driver\Driver1")
objuft.Test.run
objuft.Test.close
objuft.quit
Set objuft=nothing