from win32com import client

def Foo():
    # Connect to our COM-object
    lib = client.Dispatch('MyCOMobject')
    # Use ArrayList is a proper type of argument for Dot-method of COM-object
    values = client.Dispatch('System.Collections.ArrayList')
    values.Add(10)
    values.Add(145)
    values.Add(0.3)
    # Call Dot-method
    return lib.Dot(values)

if __name__ == "__main__":
    print(Foo())
