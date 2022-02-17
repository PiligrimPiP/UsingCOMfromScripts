//Connect to our COM-object
var lib = new ActiveXObject("MyCOMobject")

//Connect to System.Collections.ArrayList. We must send ArrayList to 'Dot' method of COM-object
var values = new ActiveXObject("System.Collections.ArrayList")
values.Add(10)
values.Add(0.3)
values.Add(145)

// Call for Dot-method
var response = lib.Dot((values))
