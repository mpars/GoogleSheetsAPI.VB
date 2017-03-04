Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks


Public Class Connect

    Shared Scopes As String() = {SheetsService.Scope.SpreadsheetsReadonly}
    Shared ApplicationName As String = "Google Sheets API VB.NET Quickstart"

Private Sub GetDataFromGoogleSheets()

        Dim credential As UserCredential
        Using stream = New FileStream("client_secret.json", FileMode.Open, FileAccess.Read)
            Dim credPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal)
            credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.vb.quickstart.json")

            ' Change "user" to your username

            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, New FileDataStore(credPath, True)).Result
            Console.WriteLine(Convert.ToString("Credential file saved to: ") & credPath)
        End Using

        Dim service = New SheetsService(New BaseClientService.Initializer() With {.HttpClientInitializer = credential, .ApplicationName = ApplicationName})


        ' ID of google sheet from url

        Dim spreadsheetId As [String] = "googlesheetID"

        ' Set the range for Data to get

        Dim range As [String] = "A2:E"

        Dim request As SpreadsheetsResource.ValuesResource.GetRequest = service.Spreadsheets.Values.[Get](spreadsheetId, range)
        Dim response As ValueRange = request.Execute()
        Dim values As IList(Of IList(Of [Object])) = response.Values

        ' If there is Data in the range

        If values IsNot Nothing AndAlso values.Count > 0 Then

            For Each RowItem In values
                Console.WriteLine(RowItem(0) + ":" + RowItem(1) + ":" + RowItem(2) + ":" + RowItem(3))
            Next
            
            Console.WriteLine("No more data")
            
        ' If there is No Data in the range
        Else
            Console.WriteLine("No data available")
        End If

    End Sub
End Class
