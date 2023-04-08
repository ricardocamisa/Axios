# Axios Class for VBA
The Axios class is a VBA project that allows you to make HTTP requests with an easy-to-use interface similar to the Axios library in JavaScript. It provides methods for GET, POST, PUT, DELETE, and PATCH requests and allows you to configure headers, data, and URLs.

## Installation

- Import the `Axios` class module into your VBA project, [**click here**](https://github.com/ricardocamisa/axios).
- Import the `ConvertJson` module from the [**VBA-JSON**](https://github.com/VBA-tools/VBA-JSON) project into your VBA project.
- To use the Axios class, you need to add the following references to your VBA project:

    `Microsoft Scripting Runtime`
    `Microsoft XML, v6.0 (or a later version)`

To add these references, open the VBA editor (Alt + F11), go to the "Tools" menu, select "References...", and check the boxes next to ``Microsoft Scripting Runtime`` and ``Microsoft XML, v6.0 (or a later version)``.

## Usage

### Create a request

To create a new `Axios` request, use the `configAxios` method. This method will return a new Axios object with the specified configuration.

```vba
' Set the authorization token
Public Token As String

Sub Exemple()
    ' Define the request headers
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers("Content-Type") = "application/json"
    headers("Authorization") = "Bearer " & Token

    ' Define the request data
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    data("email") = "admin@gmail.com"
    data("password") = "123"

    ' Create the request configuration
    Dim req As Axios
    Set req = New Axios

    With req
        .baseURL = "http://localhost:3030/api"
        .url = "/clients"
        .method = eGET
        Set .headers = headers
        Set .data = data
        Debug.Print .Send(.configAxios)
    End With
End Sub
```

### Credits
This project is based on the [**axios**](https://github.com/axios/axios) library for JavaScript.
