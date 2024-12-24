'MIT License

'Copyright (c) 2023 Ricardo Camisa

'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

Option Explicit

Private Type TAxios
    baseUrl As String
    url As String
    data As Object
    headers As Object
    method As Integer
End Type

Private this As TAxios

Public Function ConfigAxios() As Axios
    Set ConfigAxios = New Axios
    ConfigAxios.baseUrl = this.baseUrl
    ConfigAxios.url = this.url
    Set ConfigAxios.data = this.data
    Set ConfigAxios.headers = this.headers
    ConfigAxios.method = this.method
End Function


Public Function Send(config As Axios)
    Dim data As String
    data = ConvertToJson(config.data, Whitespace:=2)
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    Dim arr: arr = Array("GET", "POST", "PUT", "DELETE", "PATCH")
    
    On Error GoTo ErrorHandler
    Dim header As Variant
    
    With http
        .Open arr(config.method - 1), config.baseUrl & config.url, False
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

Function GET_(url As String, Optional params As String)
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers("Content-Type") = "application/json"
    headers("Authorization") = "Bearer " & Token
    
    With Me
        .baseUrl = baseUrl & url & "?" & params
        .method = 1
        .headers = headers
        GET_ = .Send(.ConfigAxios)
    End With
End Function

Function POST_(url As String, data As Object)
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers("Content-Type") = "application/json"
    headers("Authorization") = "Bearer " & Token
    
    With Me
        .baseUrl = baseUrl & url
        .method = 2
        .data = data
        .headers = headers
        POST_ = .Send(.ConfigAxios)
    End With
End Function

Function PACTH_(url As String, data As Object)
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers("Content-Type") = "application/json"
    headers("Authorization") = "Bearer " & Token
    
    With Me
        .baseUrl = baseUrl & url
        .method = 5
        .data = data
        .headers = headers
        PACTH_ = .Send(.ConfigAxios)
    End With
End Function
Function DELETE_(url As String)
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers("Content-Type") = "application/json"
    headers("Authorization") = "Bearer " & Token
    
    With Me
        .baseUrl = baseUrl & url
        .method = 4
        .headers = headers
        DELETE_ = .Send(.ConfigAxios)
    End With
End Function
