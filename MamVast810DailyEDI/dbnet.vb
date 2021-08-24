Imports dbClassNET

Public Class dbnet
  Public Const MFIPROD As String = "mfiprod"
  Public Const MFIDEV As String = "mfidev"
  Public Const PROD As String = "R34FILES"
  Public Const DEV As String = "TESTDATA"
  Public Shared library As String = String.Empty

  Public Const dsn As String = MFIPROD

  Private Shared Function CreateConnection() As iDB2Connection
    Dim dbc As New dbClassNET.dbClassNET

    If dsn = MFIPROD Then
      dbc.setDefaultMfiprodConnection()
      library = PROD
    Else
      dbc.setDefaultMfidevConnection()
      library = DEV
    End If

    Dim cn As iDB2Connection = dbc.getConnection
    cn.Open()
    Return cn
  End Function

  Public Shared Function MamVast810DataPull() As DataSet
    Dim dataset As New DataSet, dataAdapter As New iDB2DataAdapter
    Dim conn As iDB2Connection = CreateConnection()

    Try
      Dim sqlcommand As New iDB2Command("CALL " & library & ".SP_MAMVAST_NATIONAL_ACCOUNTS_EDI(@ACCOUNT)", conn)
      sqlcommand.CommandTimeout = 300
      sqlcommand.Parameters.Add("@ACCOUNT", iDB2DbType.iDB2VarChar).Value = "0422498"
      dataAdapter.SelectCommand = sqlcommand
      dataAdapter.Fill(dataset)
    Catch ex As Exception
      Email.LogError(ex)
      Throw New Exception(ex.Message)
    Finally
      conn.Close()
    End Try

    Return dataset
  End Function

  Public Shared Function MamVast810DataPullOneOrder(ByVal account As String, ByVal invoice As Integer) As DataSet
    Dim dataset As New DataSet, dataAdapter As New iDB2DataAdapter
    Dim conn As iDB2Connection = CreateConnection()

    Try
      Dim sqlcommand As New iDB2Command("CALL " & library & ".SP_MAMVAST_NATIONAL_ACCOUNTS_EDI_ONE_ORDER(@ACCOUNT, @INVOICE)", conn)

      sqlcommand.CommandTimeout = 300
      sqlcommand.Parameters.Add("@ACCOUNT", iDB2DbType.iDB2VarChar).Value = account
      sqlcommand.Parameters.Add("@INVOICE", iDB2DbType.iDB2Numeric).Value = invoice
      dataAdapter.SelectCommand = sqlcommand
      dataAdapter.Fill(dataset)
    Catch ex As Exception
      Email.LogError(ex)
      Throw New Exception(ex.Message)
    Finally
      conn.Close()
    End Try

    Return dataset
  End Function

  Public Shared Function CreateEDIBatchNo() As Boolean
    Dim dataset As New DataSet, dataAdapter As New iDB2DataAdapter
    Dim conn As iDB2Connection = CreateConnection()
    Dim success As Boolean = False

    Try
      Dim sqlcommand As New iDB2Command("CALL " & library & ".SP_MAMVAST_BATCH_CREATE()", conn)

      sqlcommand.CommandTimeout = 300
      sqlcommand.ExecuteNonQuery()
      success = True
    Catch ex As Exception
      Email.LogError(ex)
      Throw New Exception(ex.Message)
    Finally
      conn.Close()
    End Try
    Return success
  End Function

  Public Shared Function GetEDIBatchNo() As DataSet
    Dim dataset As New DataSet, dataAdapter As New iDB2DataAdapter
    Dim conn As iDB2Connection = CreateConnection()

    Try
      Dim sqlcommand As New iDB2Command("CALL " & library & ".SP_MAMVAST_BATCH_GET ()", conn)

      sqlcommand.CommandTimeout = 300
      dataAdapter.SelectCommand = sqlcommand
      dataAdapter.Fill(dataset)
    Catch ex As Exception
      Email.LogError(ex)
      Throw New Exception(ex.Message)
    Finally
      conn.Close()
    End Try

    Return dataset
  End Function
End Class
