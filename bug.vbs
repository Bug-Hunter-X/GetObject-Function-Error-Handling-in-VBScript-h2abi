Function GetObject() is used to create an instance of an object. If the object is not found then it returns an error. This can cause unexpected behavior in the application. For example, if you use this function to get the object from the registry and if the object is not found then you will get an error. The code will not handle this error correctly and will cause unexpected behavior.