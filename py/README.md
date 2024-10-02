# Call Excel add-ins from Python

The `py.xll` add-in calls 

Use Python [`ctypes`](https://docs.python.org/3/library/ctypes.html)
to load and specify the signature of `xll::AddIn` functions.

A `xll::AddIn` object contains all the information required to 
make the corresponding function callable from Python.


```
C:\[dir]> py
import <addin>.py

```

## Python	

Calling functions in an xll from Python.	 
Must have `xlcall32.dll` in the same directory as the xll.  

```
from ctypes import *
lib = WinDLL(<full path to xll>)
func = lib.func
func.restype = c_xxx
func.argtypes = [c_xxx, ...]	
print(func(xxx, ...))
```

Use `dumpbin /exports` to see the functions exported by the xll.  
Use `getattr(lib, 'func')` to get the function from the library if func has funny characters.

```

```

```
class OPER(Structure):
	_fields_ = [("xltype", c_int), ("val", c_double)]
	def __init__(self, val):
		self.xltype = 0
		self.val = val
	def __str__(self):
		return str(self.val)
	def __repr__(self):
		return str(self.val)
```
