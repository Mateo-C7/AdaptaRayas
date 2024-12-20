﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Este código fue generado por una herramienta.
'     Versión de runtime:4.0.30319.42000
'
'     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
'     se vuelve a generar el código.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'Microsoft.VSDesigner generó automáticamente este código fuente, versión=4.0.30319.42000.
'
Namespace WebReference1
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="WSUNOEESoap", [Namespace]:="http://tempuri.org/")>  _
    Partial Public Class WSUNOEE
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private CrearConexionXMLOperationCompleted As System.Threading.SendOrPostCallback
        
        Private EjecutarConsultaXMLOperationCompleted As System.Threading.SendOrPostCallback
        
        Private LeerEsquemaParametrosOperationCompleted As System.Threading.SendOrPostCallback
        
        Private ImportarXMLOperationCompleted As System.Threading.SendOrPostCallback
        
        Private SiesaWEBContabilizarOperationCompleted As System.Threading.SendOrPostCallback
        
        Private useDefaultCredentialsSetExplicitly As Boolean
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = Global.CreaRQI.My.MySettings.Default.CreaRQI1_WebReference1_WSUNOEE
            If (Me.IsLocalFileSystemWebService(Me.Url) = true) Then
                Me.UseDefaultCredentials = true
                Me.useDefaultCredentialsSetExplicitly = false
            Else
                Me.useDefaultCredentialsSetExplicitly = true
            End If
        End Sub
        
        Public Shadows Property Url() As String
            Get
                Return MyBase.Url
            End Get
            Set
                If (((Me.IsLocalFileSystemWebService(MyBase.Url) = true)  _
                            AndAlso (Me.useDefaultCredentialsSetExplicitly = false))  _
                            AndAlso (Me.IsLocalFileSystemWebService(value) = false)) Then
                    MyBase.UseDefaultCredentials = false
                End If
                MyBase.Url = value
            End Set
        End Property
        
        Public Shadows Property UseDefaultCredentials() As Boolean
            Get
                Return MyBase.UseDefaultCredentials
            End Get
            Set
                MyBase.UseDefaultCredentials = value
                Me.useDefaultCredentialsSetExplicitly = true
            End Set
        End Property
        
        '''<remarks/>
        Public Event CrearConexionXMLCompleted As CrearConexionXMLCompletedEventHandler
        
        '''<remarks/>
        Public Event EjecutarConsultaXMLCompleted As EjecutarConsultaXMLCompletedEventHandler
        
        '''<remarks/>
        Public Event LeerEsquemaParametrosCompleted As LeerEsquemaParametrosCompletedEventHandler
        
        '''<remarks/>
        Public Event ImportarXMLCompleted As ImportarXMLCompletedEventHandler
        
        '''<remarks/>
        Public Event SiesaWEBContabilizarCompleted As SiesaWEBContabilizarCompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/CrearConexionXML", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function CrearConexionXML(ByVal pvstrxmlConexion As String) As Boolean
            Dim results() As Object = Me.Invoke("CrearConexionXML", New Object() {pvstrxmlConexion})
            Return CType(results(0),Boolean)
        End Function
        
        '''<remarks/>
        Public Overloads Sub CrearConexionXMLAsync(ByVal pvstrxmlConexion As String)
            Me.CrearConexionXMLAsync(pvstrxmlConexion, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub CrearConexionXMLAsync(ByVal pvstrxmlConexion As String, ByVal userState As Object)
            If (Me.CrearConexionXMLOperationCompleted Is Nothing) Then
                Me.CrearConexionXMLOperationCompleted = AddressOf Me.OnCrearConexionXMLOperationCompleted
            End If
            Me.InvokeAsync("CrearConexionXML", New Object() {pvstrxmlConexion}, Me.CrearConexionXMLOperationCompleted, userState)
        End Sub
        
        Private Sub OnCrearConexionXMLOperationCompleted(ByVal arg As Object)
            If (Not (Me.CrearConexionXMLCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent CrearConexionXMLCompleted(Me, New CrearConexionXMLCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/EjecutarConsultaXML", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function EjecutarConsultaXML(ByVal pvstrxmlParametros As String) As System.Data.DataSet
            Dim results() As Object = Me.Invoke("EjecutarConsultaXML", New Object() {pvstrxmlParametros})
            Return CType(results(0),System.Data.DataSet)
        End Function
        
        '''<remarks/>
        Public Overloads Sub EjecutarConsultaXMLAsync(ByVal pvstrxmlParametros As String)
            Me.EjecutarConsultaXMLAsync(pvstrxmlParametros, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub EjecutarConsultaXMLAsync(ByVal pvstrxmlParametros As String, ByVal userState As Object)
            If (Me.EjecutarConsultaXMLOperationCompleted Is Nothing) Then
                Me.EjecutarConsultaXMLOperationCompleted = AddressOf Me.OnEjecutarConsultaXMLOperationCompleted
            End If
            Me.InvokeAsync("EjecutarConsultaXML", New Object() {pvstrxmlParametros}, Me.EjecutarConsultaXMLOperationCompleted, userState)
        End Sub
        
        Private Sub OnEjecutarConsultaXMLOperationCompleted(ByVal arg As Object)
            If (Not (Me.EjecutarConsultaXMLCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent EjecutarConsultaXMLCompleted(Me, New EjecutarConsultaXMLCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/LeerEsquemaParametros", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function LeerEsquemaParametros(ByVal pvstrxmlParametros As String) As String
            Dim results() As Object = Me.Invoke("LeerEsquemaParametros", New Object() {pvstrxmlParametros})
            Return CType(results(0),String)
        End Function
        
        '''<remarks/>
        Public Overloads Sub LeerEsquemaParametrosAsync(ByVal pvstrxmlParametros As String)
            Me.LeerEsquemaParametrosAsync(pvstrxmlParametros, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub LeerEsquemaParametrosAsync(ByVal pvstrxmlParametros As String, ByVal userState As Object)
            If (Me.LeerEsquemaParametrosOperationCompleted Is Nothing) Then
                Me.LeerEsquemaParametrosOperationCompleted = AddressOf Me.OnLeerEsquemaParametrosOperationCompleted
            End If
            Me.InvokeAsync("LeerEsquemaParametros", New Object() {pvstrxmlParametros}, Me.LeerEsquemaParametrosOperationCompleted, userState)
        End Sub
        
        Private Sub OnLeerEsquemaParametrosOperationCompleted(ByVal arg As Object)
            If (Not (Me.LeerEsquemaParametrosCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent LeerEsquemaParametrosCompleted(Me, New LeerEsquemaParametrosCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/ImportarXML", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function ImportarXML(ByVal pvstrDatos As String, ByRef printTipoError As Short) As System.Data.DataSet
            Dim results() As Object = Me.Invoke("ImportarXML", New Object() {pvstrDatos, printTipoError})
            printTipoError = CType(results(1),Short)
            Return CType(results(0),System.Data.DataSet)
        End Function
        
        '''<remarks/>
        Public Overloads Sub ImportarXMLAsync(ByVal pvstrDatos As String, ByVal printTipoError As Short)
            Me.ImportarXMLAsync(pvstrDatos, printTipoError, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub ImportarXMLAsync(ByVal pvstrDatos As String, ByVal printTipoError As Short, ByVal userState As Object)
            If (Me.ImportarXMLOperationCompleted Is Nothing) Then
                Me.ImportarXMLOperationCompleted = AddressOf Me.OnImportarXMLOperationCompleted
            End If
            Me.InvokeAsync("ImportarXML", New Object() {pvstrDatos, printTipoError}, Me.ImportarXMLOperationCompleted, userState)
        End Sub
        
        Private Sub OnImportarXMLOperationCompleted(ByVal arg As Object)
            If (Not (Me.ImportarXMLCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent ImportarXMLCompleted(Me, New ImportarXMLCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SiesaWEBContabilizar", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function SiesaWEBContabilizar(ByVal pvstrParametros As String) As Short
            Dim results() As Object = Me.Invoke("SiesaWEBContabilizar", New Object() {pvstrParametros})
            Return CType(results(0),Short)
        End Function
        
        '''<remarks/>
        Public Overloads Sub SiesaWEBContabilizarAsync(ByVal pvstrParametros As String)
            Me.SiesaWEBContabilizarAsync(pvstrParametros, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub SiesaWEBContabilizarAsync(ByVal pvstrParametros As String, ByVal userState As Object)
            If (Me.SiesaWEBContabilizarOperationCompleted Is Nothing) Then
                Me.SiesaWEBContabilizarOperationCompleted = AddressOf Me.OnSiesaWEBContabilizarOperationCompleted
            End If
            Me.InvokeAsync("SiesaWEBContabilizar", New Object() {pvstrParametros}, Me.SiesaWEBContabilizarOperationCompleted, userState)
        End Sub
        
        Private Sub OnSiesaWEBContabilizarOperationCompleted(ByVal arg As Object)
            If (Not (Me.SiesaWEBContabilizarCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent SiesaWEBContabilizarCompleted(Me, New SiesaWEBContabilizarCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        Public Shadows Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub
        
        Private Function IsLocalFileSystemWebService(ByVal url As String) As Boolean
            If ((url Is Nothing)  _
                        OrElse (url Is String.Empty)) Then
                Return false
            End If
            Dim wsUri As System.Uri = New System.Uri(url)
            If ((wsUri.Port >= 1024)  _
                        AndAlso (String.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) = 0)) Then
                Return true
            End If
            Return false
        End Function
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub CrearConexionXMLCompletedEventHandler(ByVal sender As Object, ByVal e As CrearConexionXMLCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class CrearConexionXMLCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As Boolean
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),Boolean)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub EjecutarConsultaXMLCompletedEventHandler(ByVal sender As Object, ByVal e As EjecutarConsultaXMLCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class EjecutarConsultaXMLCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataSet
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataSet)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub LeerEsquemaParametrosCompletedEventHandler(ByVal sender As Object, ByVal e As LeerEsquemaParametrosCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class LeerEsquemaParametrosCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As String
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),String)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub ImportarXMLCompletedEventHandler(ByVal sender As Object, ByVal e As ImportarXMLCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class ImportarXMLCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataSet
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataSet)
            End Get
        End Property
        
        '''<remarks/>
        Public ReadOnly Property printTipoError() As Short
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(1),Short)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub SiesaWEBContabilizarCompletedEventHandler(ByVal sender As Object, ByVal e As SiesaWEBContabilizarCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class SiesaWEBContabilizarCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As Short
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),Short)
            End Get
        End Property
    End Class
End Namespace
