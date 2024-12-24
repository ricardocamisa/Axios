Option Explicit

Private Type IInstance
    baseUrl As String
    timeout As Long
    headers As Object
    data As String
End Type

Private this As IInstance

' Constructor
Private Sub Class_Initialize()
    ' Inicializa a instância com valores padrão
    With this
        .baseUrl = "http://127.0.0.1:8090" ' URL base da API
        .timeout = 1000
        Set .headers = CreateObject("Scripting.Dictionary")
        .headers("Content-Type") = "application/json"
        .headers("Authorization") = "Bearer " & Token ' Substitua pelo seu token
    End With
End Sub

' Configura uma nova instância personalizada
Public Sub create(baseUrl As String, Optional timeout As Long = 1000, Optional headers As Object = Nothing)
    With this
        .baseUrl = baseUrl
        .timeout = timeout
        If Not headers Is Nothing Then
            Set .headers = headers
        End If
    End With
End Sub

' Função para enviar requisições
Private Function request(method As String, endpoint As String, Optional data As Object = Nothing) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    this.data = ConvertToJson(data, Whitespace:=2)
    
    Dim url As String
    url = this.baseUrl & endpoint
    
    ' Configura o tempo limite
    http.setTimeouts this.timeout, this.timeout, this.timeout, this.timeout
    
    ' Abre a requisição
    http.Open method, url, False
    
    ' Configura os cabeçalhos
    Dim key As Variant
    For Each key In this.headers.Keys
        http.setRequestHeader key, this.headers(key)
    Next key
    
    ' Envia a requisição
    On Error GoTo RequestError
    http.Send this.data
    On Error GoTo 0
    
    ' Retorna a resposta
    request = http.responseText
    Exit Function
    
RequestError:
    request = "Erro: " & Err.Description
End Function

' Métodos auxiliares para atalhos de requisição
Public Function get_(endpoint As String) As String
    get_ = request("GET", endpoint)
End Function

Public Function post_(endpoint As String, Optional data As Object) As String
    post_ = request("POST", endpoint, data)
End Function

Public Function put_(endpoint As String, Optional data As Object) As String
    put_ = request("PUT", endpoint, data)
End Function

Public Function delete_(endpoint As String) As String
    delete_ = request("DELETE", endpoint)
End Function
