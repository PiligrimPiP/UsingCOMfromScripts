from win32com import client

def Foo():
    lib = client.Dispatch('MyCOMobject')
    values = client.Dispatch('System.Collections.ArrayList')
    values.Add(10)
    values.Add(2.3)
    return lib.dot(values)

if __name__ == "__main__":
    print(Foo())