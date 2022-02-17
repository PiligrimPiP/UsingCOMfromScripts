Function Foo()
	' Simply connect to any COM using CreateObject-instance
	set lib = CreateObject("MyCOMobject")
	set values = CreateObject("System.Collections.ArrayList")
	values.Add 10
	values.Add 2
	' Call Dot-method of DLL and return value
	Foo = lib.Dot((values))

End Function
