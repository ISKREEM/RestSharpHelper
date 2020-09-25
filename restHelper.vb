Imports System.Net
Imports RestSharp

Module Main

    Sub Main()
        Console.WriteLine("...Iniciando request...")
        'Se instancia la clase auxiliar generica
        Dim warzone As Warzone = New Warzone("http://localhost/api/v1.0/")
        'Se realiza el request indicado sobre el recurso deseado indicando el verbo de la operacion asi como demas parametros
        Dim result = warzone.GetResult(Of RestEntity)("method", Method.GET)
        'Se valida el resultado obtenido
        Try
            warzone.ValidateResult(result)
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
        Console.WriteLine("...Finalizado...")
        Console.ReadKey()
    End Sub

End Module

Class Warzone
    Private Property _execute As RestClient

    ''' <summary>
    ''' Constructor de la clase que recibe la url base del servicio rest
    ''' </summary>
    ''' <param name="resource">url base, ej. http://localhost/api/...</param>
    Public Sub New(ByVal resource As String)
        'Se crea la instancia del cliente
        Me._execute = New RestClient(resource)
    End Sub

    ''' <summary>
    ''' Funcion generica para ejecutar un request implementando RestSharp
    ''' </summary>
    ''' <typeparam name="T">Clase o Entidad para deserializar el reponse</typeparam>
    ''' <param name="resource">metodo del recurso rest a usar</param>
    ''' <param name="method">Verbo HTTP - GET, POST, PUT, DELETE</param>
    ''' <param name="headers">Headers del request</param>
    ''' <param name="parameters">Parametros del request</param>
    ''' <param name="body">Objeto a serializar para casos con JsonBody</param>
    ''' <param name="isJsonParams">Indicador para request con JsonBody ej. {'key': 'value'}</param>
    ''' <param name="isQueryParams">Indicador para request con Query Parameters ej. ...?param=value&param2=value</param>
    ''' <returns>Tupla HTTP Code(200, 204, 409, 500) y Objeto indicado(T)</returns>
    Public Function GetResult(Of T As New)(ByVal resource As String, ByVal method As Method,
                                           Optional ByVal headers As Dictionary(Of String, String) = Nothing,
                                           Optional ByVal parameters As Dictionary(Of String, Object) = Nothing,
                                           Optional ByVal body As Object = Nothing,
                                           Optional ByVal isJsonParams As Boolean = False,
                                           Optional ByVal isQueryParams As Boolean = False,
                                           Optional ByVal isMultipart As Boolean = False) As Tuple(Of HttpStatusCode, T, String)
        'Inicializamos el request con el metodo del api y el verbo Http
        Dim request = New RestRequest(resource, method)
        'Creamos la variable para almacenar el response
        Dim response As IRestResponse(Of T)

        'En caso de utilizar hearders los agregamos
        If Not IsNothing(headers) Then
            For Each item In headers
                request.AddHeader(item.Key, item.Value)
            Next
        End If

        'Si tenemos parametros para el request los agregamos
        If Not IsNothing(parameters) Then
            For Each item In parameters
                request.AddParameter(item.Key, item.Value)
            Next
        End If

        'Si el request es de tipo json lo agregamos
        If isJsonParams And body IsNot Nothing Then
            request.AddJsonBody(body)
        End If

        'Si el request requiere parmetros por url los agregamos
        If isQueryParams And Not IsNothing(parameters) Then
            For Each item In parameters
                request.AddQueryParameter(item.Key, item.Value)
            Next
        End If

        'Si el request es tiene formato multipart/form
        request.AlwaysMultipartFormData = isMultipart

        Try
            'Se realiza la peticion
            response = _execute.Execute(Of T)(request)
            'Se valida si existe algun tipo de error
            If Not String.IsNullOrWhiteSpace(response.ErrorMessage) Then
                Throw New Exception(response.ErrorMessage)
            End If
        Catch ex As Exception
            'En caso se fallo se indica el error con el tercer parametro
            Return Tuple.Create(HttpStatusCode.InternalServerError, New T, $"{ex.Message} {vbCrLf} {ex.StackTrace}")
        End Try
        'Si no hubo problema alguno con la peticion se retorna el resultado
        Return Tuple.Create(response.StatusCode, response.Data, String.Empty)
    End Function

    ''' <summary>
    ''' Funcion para validar los codigos http recibidos
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="result"></param>
    ''' <returns></returns>
    Public Function ValidateResult(Of T As New)(ByVal result As Tuple(Of HttpStatusCode, T, String)) As T
        If result.Item1 = 200 Then
            Return result.Item2
        ElseIf result.Item1 = 204 Then
            Return New T
        Else
            'Una de muchas...
            Throw New Exception(IIf(result.Item1 = 400, "Bad request", result.Item3))
        End If
    End Function

End Class

''' <summary>
''' Entidad que sera implementada para la deserializacion del response
''' Las propiedades pueden variar de reponse a response, las propiedades deben tener el mismo nombre que las del response
''' </summary>
Class RestEntity
    Public Property [error] As Boolean
    Public Property message As String
    Public Property data As Dota2 'data puede ser de cualquier tipo por ejemplo: List(Of Dota2), Dicctionary(Of Dota2), Dota2, etc
    Public Property status As String
End Class

Class Dota2
    Public Property rank As Int32
    Public Property medal As String
    Public Property nickname As String
End Class
