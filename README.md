# Axios Class for VBA
The Axios class is a VBA project that allows you to make HTTP requests with an easy-to-use interface similar to the Axios library in JavaScript. It provides methods for GET, POST, PUT, DELETE, and PATCH requests and allows you to configure headers, data, and URLs.

## Installation

- Import the `Axios` class module into your VBA project, [**click here**](https://github.com/ricardocamisa/Axios/blob/main/Axios.cls).
- Import the `ConvertJson` module from the [**VBA-JSON**](https://github.com/VBA-tools/VBA-JSON) project into your VBA project.
- To use the Axios class, you need to add the following references to your VBA project:

    `Microsoft Scripting Runtime`
    `Microsoft XML, v6.0 (or a later version)`

- To add these references, open the VBA editor (Alt + F11), go to the "Tools" menu, select "References...", and check the boxes next to ``Microsoft Scripting Runtime`` and ``Microsoft XML, v6.0 (or a later version)``.

## Usage

### Create a request

To create a new `Axios` request, use the `configAxios` method. This method will return a new Axios object with the specified configuration.

```vba
' Set the authorization token
Public Token As String

Sub Exemple(identity as string, password as string)
     Dim Axios As New Axios
     Dim props As Object

    Dim requestBody As Object
    Dim response As String
    
    Set requestBody = CreateObject("Scripting.Dictionary")
    With requestBody
        .Add "identity", identity
        .Add "password", password
    End With

    If InStr(response, "Erro: Não foi possível estabelecer uma ligação ao servidor") > 0 Then
        MsgBox "Erro: Não foi possível estabelecer uma ligação ao servidor"
    Else
        Dim Json As Object
        Set Json = JsonConverter.ParseJson(response)
        
        ' Verifica se o campo "token" existe antes de acessar
        If Not Json Is Nothing And Json.Exists("token") Then
            AccessToken = Json("token")
            Debug.Print Json("token")
            MsgBox "Seja bem-vindo de volta " & Json("record")("name"), vbInformation, "Sucesso"
        Else
            If Not Json Is Nothing And Json.Exists("message") Then
                MsgBox Json("message"), vbCritical, "Erro de autenticação!"
            Else
                MsgBox "O campo 'token' não foi encontrado na resposta.", vbExclamation, "Erro!"
            End If
        End If
    End If
End Sub
```

### Credits
This project is based on the [**axios**](https://github.com/axios/axios) library for JavaScript.
