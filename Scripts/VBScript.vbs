Function Foo()
	
	set lib =CreateObject("MyCOMobject")
	set values = CreateObject("System.Collections.ArrayList")
	values.Add 10
	values.Add 2

	Foo = lib.Dot((values))

End Function