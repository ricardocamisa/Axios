'MIT License

'Copyright (c) 2023 Ricardo Camisa

'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

Option Explicit

Public http As Object

Enum HttpMethod
    eGET = 1
    ePOST = 2
    ePUT = 3
    eDELETE = 4
    ePATCH = 5
End Enum

Private Type AxiosConfig
    baseURL As String
    url As String
    data As Object
    headers As Object
    method As HttpMethod
End Type

Private this As AxiosConfig

Public Property Get baseURL() As String
    baseURL = this.baseURL
End Property

Public Property Let baseURL(value As String)
    this.baseURL = value
End Property

Public Property Get url() As String
    url = this.url
End Property

Public Property Let url(value As String)
    this.url = value
End Property

Public Property Get data() As Object
    Set data = this.data
End Property

Public Property Set data(value As Object)
    Set this.data = value
End Property

Public Property Get headers() As Object
    Set headers = this.headers
End Property

Public Property Set headers(value As Object)
    Set this.headers = value
End Property

Public Property Get method() As HttpMethod
    method = this.method
End Property

Public Property Let method(value As HttpMethod)
    this.method = value
End Property

Public Function configAxios() As Axios
    Set configAxios = New Axios
    configAxios.baseURL = this.baseURL
    configAxios.url = this.url
    Set configAxios.data = this.data
    Set configAxios.headers = this.headers
    configAxios.method = this.method
End Function


Public Function Send(config As Axios)
    Dim data As String
    data = ConvertToJson(config.data, Whitespace:=2)
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    Dim arr: arr = Array("GET", "POST", "PUT", "DELETE", "PATCH")
    
        
    On Error GoTo ErrorHandler
    Dim header As Variant
    
    With http
        .Open arr(config.method - 1), config.baseURL & config.url, False
        For Each header In config.headers
            .setRequestHeader header, config.headers(header)
        Next header
        .Send data
        Send = .responseText
    End With
    
    Exit Function
    
ErrorHandler:
    Send = "Erro ao enviar solicitação: " & Err.Description
End Function
