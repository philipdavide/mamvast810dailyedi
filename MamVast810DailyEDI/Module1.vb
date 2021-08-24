Imports System.Net
Imports System.Text

Module Module1
  Public Const batchNoPadCount As Integer = 9
  Public Const qualifierID As String = "ZZ"
  Public Const senderID As String = "MAXFINKELSTEIN"

  Public Const ftp_mamvast_upload_url_dev As String = "ftp://Ftpaws.maxfinkelstein.com"
  Public Const ftp_mamvast_upload_login_dev As String = "mamm"
  Public Const ftp_mamvast_upload_pass_dev As String = "nat2840!"
  Public Const send_to_ftp As Boolean = True

  Sub Main()
    Console.WriteLine("MamVast 810 Daily EDI")
    Console.WriteLine()
    Console.WriteLine("Pulling transaction records from the previous day.")

    Dim dataset As New DataSet
    Dim batchNo As Integer = 0
    Try
      'for testing one invoice as a time
      '  dataset = dbnet.MamVast810DataPullOneOrder("0422498", 877803)
      'for multiple invoices in one file for one acccount (live version)
      dataset = dbnet.MamVast810DataPull()
      Dim batchNoMade As Boolean = dbnet.CreateEDIBatchNo()

      If Not batchNoMade Then
        Email.LogError(New Exception("Unable to create batch number"))
        Console.WriteLine("Error Creating a Batch Number")
        Console.Read()
        Exit Sub
      End If

      Dim batchDataset As DataSet = dbnet.GetEDIBatchNo()
      batchNo = batchDataset.Tables(0).Rows(0).Item("MAMBATCH")
    Catch ex As Exception
      Email.LogError(ex)
      Console.WriteLine(ex.Message)
      Console.Read()
      Exit Sub
    End Try

    Dim ediStringBuilder As New System.Text.StringBuilder
    Dim passheadDatatable As DataTable = dataset.Tables(0)
    Dim shippingDatatable As DataTable = dataset.Tables(1)
    Dim billingTable As DataTable = dataset.Tables(2)
    Dim passdetailDatatable As DataTable = dataset.Tables(3)

    ' Dim totolBatchCharacter As Integer = 9
    Dim transactionCount As String = batchNo.ToString().PadRight(batchNoPadCount, "0") '1
    Dim account As String = String.Empty

    Try
      ' the * is a separator here

      For Each headerRow As DataRow In passheadDatatable.Rows
        'check if the account as already used for the EDI
        If account = headerRow.Item("XPHCUST") Then Continue For

        account = headerRow.Item("XPHCUST")

        'build Interchange Control Header
        'Authorization Information Qualifier
        'Authorization Information
        'Security Information Qualifier
        'Security Information
        'Interchange ID Qualifier
        'Interchange Sender ID
        'Interchange ID Qualifier
        'Interchange Receiver ID
        'Interchange Date
        'Interchange Time
        'Interchange Control Standards Identifier
        'Interchange Control Version Number
        'Interchange Control Number, must be incremented, database table needed?, padding 9
        'Acknowledgment Requested
        'Test Indicator
        ediStringBuilder.AppendFormat("ISA*00*          *00*          *{0}*{1}*{0}*{2}*{3}*{4}*U*00401*{5}*0*T*>",
                                      qualifierID,
                                      senderID.PadRight(15, " "),
                                      headerRow.Item("XPHCUST").PadRight(15, " "),
                                      Now.ToString("yyMMdd"),
                                      Now.ToString("HHmm"),
                                      batchNo.ToString().PadLeft(batchNoPadCount, "0")).AppendLine("~")

        'build Functional Group Header
        'Functional Identifier Code
        'Application Sender's Code
        'Application Sender's Code
        'Application Receiver's Code
        'Date
        'Time
        'Group Control Number, must be incremented, database table needed?
        'Responsible Agency Code
        ediStringBuilder.AppendFormat("GS*IN*{0}*{1}*{2}*{3}*{4}*X*004010", senderID,
                                      headerRow.Item("XPHCUST"),
                                      Now.ToString("yyyyMMdd"),
                                      Now.ToString("HHmm"),
                                      batchNo).AppendLine("~")

        Dim headerView As DataView = passheadDatatable.DefaultView()
        headerView.RowFilter = "XPHCUST = '" & account & "'"

        Dim headerDatatable As DataTable = headerView.ToTable()

        For Each headerAccountRow As DataRow In headerDatatable.Rows
          Dim segmentCount As Integer = 0   'segments count for things like 'ST', 'BIG', 'REF', 'N1'
          transactionCount = Integer.Parse(transactionCount) + 1
          'build Transaction Set Header
          ediStringBuilder.AppendFormat("ST*810*{0}", transactionCount).AppendLine("~")    'Transaction Set Control Number
          segmentCount += 1

          'build the header for a passthrough order
          'invoice date
          'invoice no
          'print date
          'gy source doc
          'debit or credit
          ediStringBuilder.AppendFormat("BIG*{0}*{1}*{2}*{3}***{4}", headerAccountRow.Item("XPHTDTE"),
                                      headerAccountRow.Item("XPHINV"),
                                      headerAccountRow.Item("XPHPRTD"),
                                      headerAccountRow.Item("XPHCSDC"),
                                      headerAccountRow.Item("DR_CR")).AppendLine("~")

          Dim debitCredit As String = headerAccountRow.Item("DR_CR")
          segmentCount += 1

          'build references for nonsig number and order type
          ediStringBuilder.AppendFormat("REF*BAI*{0}*NON-SIG NUMBER", headerAccountRow.Item("NONSIG")).AppendLine("~")
          segmentCount += 1
          ediStringBuilder.AppendFormat("REF*TOC*{0}*TYPE", headerAccountRow.Item("PASSTYPE")).AppendLine("~")
          segmentCount += 1

          Dim shippingView As DataView = shippingDatatable.DefaultView()
          shippingView.RowFilter = "RVCUST = '" & headerAccountRow.Item("XPHCUST") & "' AND RVSHIP = '" & headerAccountRow.Item("SHIPTO") & "'"

          If shippingView.Count > 0 Then
            Dim shippingRow As DataRow = shippingView.ToTable().Rows(0)

            'build the shipping information
            ediStringBuilder.AppendFormat("N1*ST*{0}", shippingRow.Item("RVNAME")).AppendLine("~")
            segmentCount += 1
            ediStringBuilder.AppendFormat("N3*{0}", shippingRow.Item("RVADD1")).AppendLine("~")
            segmentCount += 1

            ediStringBuilder.AppendFormat("N4*{0}*{1}*{2}*US", shippingRow.Item("RVCITY"),
                                        shippingRow.Item("RVST1"),
                                        shippingRow.Item("RVZIP")).AppendLine("~")
            segmentCount += 1
          End If

          Dim billingView As DataView = billingTable.DefaultView()
          billingView.RowFilter = "RMCUST = '" & headerAccountRow.Item("XPHCUST") & "'"

          If billingView.Count > 0 Then
            Dim billingRow As DataRow = billingView.ToTable().Rows(0)

            'build the billing information
            ediStringBuilder.AppendFormat("N1*BT*{0}", billingRow.Item("RMNAME")).AppendLine("~")
            segmentCount += 1

            ediStringBuilder.AppendFormat("N3*{0}", billingRow.Item("RMADD1")).AppendLine("~")
            segmentCount += 1

            ediStringBuilder.AppendFormat("N4*{0}*{1}*{2}*US", billingRow.Item("RMCITY"),
                                        billingRow.Item("RMSTAT"),
                                        billingRow.Item("RMZIP")).AppendLine("~")
            segmentCount += 1
          End If

          ediStringBuilder.AppendFormat("ITD*05*3**{0}**{0}*******", headerAccountRow.Item("XPHTDTE")).AppendLine("~")
          segmentCount += 1

          Dim detailView As DataView = passdetailDatatable.DefaultView()
          detailView.RowFilter = "XPHCUST = '" & headerAccountRow.Item("XPHCUST") & "' AND XPDGYIN = '" & headerAccountRow.Item("XPHGYIN") & "'"
          'need another datatable to get the line item
          'data in order for the EDI
          Dim EDIItemsTable As DataTable = passdetailDatatable.Clone()
          EDIItemsTable.Columns.Add("LineNo", GetType(Integer))
          '   Console.WriteLine(headerAccountRow.Item("XPHINV"))

          If detailView.Count > 0 Then
            Dim detailTable As DataTable = detailView.ToTable()
            Dim lineitemCount As Integer = 1
            Dim loopCounter As Integer = 0
            Dim currentItemcode As String = String.Empty
            Dim currentQty As Integer = 0

            'use row filter - get all rows that have the itemcode
            Dim dataview As DataView = detailTable.DefaultView()

            'the order of the items go as follows:
            '1. items of sort type T/S
            '2. Deliver Commission
            '3. F.E.T.
            '**************************************
            'calculations go as following:
            '(Total Cost for itemcode (sort type T/S) - cash discount)
            '-
            'Total NET STATE PRICE (sort type G)
            'Divide by Total Qty for Itemcode
            '= Unit Value for Mam
            '***************************************
            'Credit for tire of an adjustment type invoice
            ' We have the unit value for the credit
            ' From that unit value we take away the adjustment replace price
            ' The handling allowance gets credited back
            ' The 2% cash discount is divided by the qty of the tire
            ' The result of this is taken away from the remaining unit value

            ' Total Cost for itemcode
            ' - Total Cost for adjustment replace price
            ' + handling allowance (sort type H)
            ' = adjustment cost for item (sort type A)
            '
            ' 2% cash discount / by the unit qty
            '	= cash discount for one unit

            '	adjustment cost for item
            '	- cash discount for one unit
            ' = final credit
            'build the detail for a passthrough order
            While loopCounter < detailTable.Rows.Count
              Dim row As DataRow = detailTable.Rows(0)
              'check for items only 
              Select Case row.Item("XPDSORT")
                Case "T", "S"

                  dataview.RowFilter = "XPDITEM = '" & row.Item("XPDITEM") & "'"

                  currentItemcode = row.Item("XPDITEM")

                  Dim itemTable As DataTable = dataview.ToTable()
                  Dim itemtotal As Double = 0.0
                  Dim qtyTotal As Integer = 0
                  'get the total quantity and itemcode credit
                  For Each itemRow As DataRow In itemTable.Rows
                    itemtotal += itemRow.Item("XPDQTY") * itemRow.Item("XPDAMT")
                    qtyTotal += itemRow.Item("XPDQTY")
                  Next

                  EDIItemsTable.Rows.Add(qtyTotal,
                                         row.Item("UNIT_MEAS"),
                                         itemtotal,
                                         row.Item("VENDOR_CODE"),
                                         row.Item("XPDSORT"),
                                         row.Item("XPDITEM"),
                                         row.Item("DESCRIPTION"),
                                         row.Item("XPHCUST"),
                                         row.Item("XPDGYIN"),
                                         lineitemCount)

                  currentQty = qtyTotal
                  lineitemCount += 1
                  'use the data view to delete all rows with the itemcode
                  While dataview.Count > 0
                    dataview.Delete(0)
                  End While

                  dataview.RowFilter = String.Empty
                Case "G"
                  'get the NET STATE PRICE specific to the itemcode
                  'find using quantity
                  dataview.RowFilter = "XPDSORT = 'G' AND XPDQTY = " & currentQty

                  'substract the total NET STATE PRICE 
                  'From the total itemcode credit
                  Dim ediView As DataView = EDIItemsTable.DefaultView()
                  ediView.RowFilter = "XPDITEM = '" & currentItemcode & "'"
                  ediView.Item(0).Item("XPDAMT") += (dataview.Item(0).Item("XPDAMT") * dataview.Item(0).Item("XPDQTY"))
                  'use the data view to delete the NET STATE PRICE
                  'for the itemcode
                  While dataview.Count > 0
                    dataview.Delete(0)
                  End While

                  dataview.RowFilter = String.Empty
                Case "D" '"DELIVERY COMMISSION"
                  'use the data view to delete records for 
                  'the FET/Delivery Comission for the correct itemcode
                  dataview.RowFilter = "XPDSORT = 'D' AND XPDAMT = '" & row.Item("XPDAMT") & "'"
                  Dim commissionRow As DataRowView = dataview.Item(0)

                  EDIItemsTable.Rows.Add(commissionRow.Item("XPDQTY"),
                                    commissionRow.Item("UNIT_MEAS"),
                                    commissionRow.Item("XPDAMT"),
                                    commissionRow.Item("VENDOR_CODE"),
                                    commissionRow.Item("XPDSORT"),
                                    commissionRow.Item("XPDITEM"),
                                    commissionRow.Item("DESCRIPTION"),
                                    commissionRow.Item("XPHCUST"),
                                    commissionRow.Item("XPDGYIN"),
                                    lineitemCount)
                  While dataview.Count > 0
                    dataview.Delete(0)
                  End While

                  dataview.RowFilter = String.Empty
                Case "F" '"FET"
                  'add the FET/Delivery Commission for the
                  'correct current itemcode
                  dataview.RowFilter = "XPDSORT = 'F' AND XPDAMT = '" & row.Item("XPDAMT") & "'"
                  Dim taxRow As DataRowView = dataview.Item(0)

                  EDIItemsTable.Rows.Add(taxRow.Item("XPDQTY"),
                                    taxRow.Item("UNIT_MEAS"),
                                    taxRow.Item("XPDAMT"),
                                    taxRow.Item("VENDOR_CODE"),
                                    taxRow.Item("XPDSORT"),
                                    taxRow.Item("XPDITEM"),
                                    taxRow.Item("DESCRIPTION"),
                                    taxRow.Item("XPHCUST"),
                                    taxRow.Item("XPDGYIN"),
                                    lineitemCount)
                
                  While dataview.Count > 0
                    dataview.Delete(0)
                  End While

                  dataview.RowFilter = String.Empty
                Case "C" 'CONSUMER DELIVERY
                  dataview.RowFilter = "XPDSORT = 'C'"
                  Dim ediView As DataView = EDIItemsTable.DefaultView()
                  ediView.RowFilter = "XPDITEM = '" & currentItemcode & "'"
                  'debit or credit 
                  If debitCredit = "CN" Then
                    ediView.Item(0).Item("XPDAMT") -= (dataview.Item(0).Item("XPDAMT") * dataview.Item(0).Item("XPDQTY"))
                  ElseIf debitCredit = "DI" Then
                    ediView.Item(0).Item("XPDAMT") += (dataview.Item(0).Item("XPDAMT") * dataview.Item(0).Item("XPDQTY"))
                  End If


                  dataview.Delete(0)
                  dataview.RowFilter = String.Empty
                Case "X"
                  dataview.Delete(0)
                Case "A"
                  dataview.RowFilter = "XPDSORT = 'A'"
                  Dim ediView As DataView = EDIItemsTable.DefaultView()
                  ediView.RowFilter = "XPDITEM = '" & currentItemcode & "'"
                  ediView.Item(0).Item("XPDAMT") += (dataview.Item(0).Item("XPDAMT") * dataview.Item(0).Item("XPDQTY"))
                  dataview.Delete(0)
                  dataview.RowFilter = String.Empty
                Case "H"
                  dataview.RowFilter = "XPDSORT = 'H'"
                  Dim ediView As DataView = EDIItemsTable.DefaultView()
                  ediView.RowFilter = "XPDITEM = '" & currentItemcode & "'"
                  ediView.Item(0).Item("XPDAMT") += (dataview.Item(0).Item("XPDAMT") * dataview.Item(0).Item("XPDQTY"))
                  dataview.Delete(0)
                  dataview.RowFilter = String.Empty
                Case Else
                  'USE 'CASE ELSE' TO GET AND CALCULATE
                  '"LESS 2% CASH DISC"
                  dataview.RowFilter = "XPDITEM LIKE '%CASH DISC%'"

                  If dataview.Count > 0 Then
                    'substract the total NET STATE PRICE 
                    'From the total itemcode credit
                    Dim ediView As DataView = EDIItemsTable.DefaultView()
                    ediView.RowFilter = "XPDITEM = '" & currentItemcode & "'"
                    ediView.Item(0).Item("XPDAMT") += (dataview.Item(0).Item("XPDAMT") * dataview.Item(0).Item("XPDQTY"))
                    ediView.Item(0).Item("XPDAMT") /= dataview.Item(0).Item("XPDQTY")
                  End If

                  dataview.RowFilter = String.Empty
                  dataview.Delete(0)
              End Select
            End While

            'loop to take the data from the EDI datatable
            'and format it into the EDI
            For Each row As DataRow In EDIItemsTable.Rows
              Select Case row.Item("XPDSORT")
                Case "T", "S"
                  Dim qty As Integer = row.Item("XPDQTY")
                  'get the itemcode and vendor this way
                  'in the cases of adjustment invoices and 
                  'selling greater than buying
                  Dim itemCode As String = row.Item("XPDITEM").split(" ")(0)
                
                  If debitCredit = "CN" Then qty *= -1

                  ediStringBuilder.AppendFormat("IT1*{0}*{1}*{2}*{3}*PE*IN*{4}*VN*{5}", row.Item("LineNo"),
                   qty,
                   row.Item("UNIT_MEAS"),
                   Double.Parse(row.Item("XPDAMT") / row.Item("XPDQTY")).ToString("F2"),
                   itemCode,
                   row.Item("VENDOR_CODE")).AppendLine("~")
                  segmentCount += 1

                  If Not IsDBNull(row.Item("DESCRIPTION")) Then
                    If row.Item("DESCRIPTION") <> String.Empty Then
                      ediStringBuilder.AppendFormat("PID*F****{0}", row.Item("DESCRIPTION")).AppendLine("~")
                      segmentCount += 1
                    End If
                  End If
                Case "D" '"DELIVERY COMMISSION"
                  ediStringBuilder.AppendFormat("SAC*A*B310***{0}**********{1}",
                                                row.Item("XPDAMT").ToString().Replace(".", String.Empty),
                                                row.Item("XPDITEM")).AppendLine("~")
                  segmentCount += 1
                Case "F" '"FEDERAL EXCISE TAX"
                  ediStringBuilder.AppendFormat("SAC*C*H670***{0}**********{1}", row.Item("XPDAMT").ToString().Replace(".", String.Empty),
                    row.Item("XPDITEM")).AppendLine("~")
                  segmentCount += 1
                Case Else


              End Select
            Next

            ediStringBuilder.AppendFormat("TDS*{0}", headerAccountRow.Item("XPHAMT").ToString().Replace(".", String.Empty)).AppendLine("~")
            segmentCount += 1

            ediStringBuilder.AppendFormat("CTT*{0}", lineitemCount - 1).AppendLine("~")
            segmentCount += 1
          End If

          segmentCount += 1
          'build Transaction Set Trailer
          ediStringBuilder.AppendFormat("SE*{0}*{1}", segmentCount,
                                        transactionCount).AppendLine("~")
        Next
        'build Functional Group Trailer
        ediStringBuilder.AppendFormat("GE*{0}*{1}", passheadDatatable.Rows.Count,
                                    batchNo).AppendLine("~")

        'build Interchange Control Trailer
        ediStringBuilder.AppendFormat("IEA*1*{0}", batchNo.ToString().PadLeft(batchNoPadCount, "0")).Append("~") 'Interchange Control Number, must be incremented, database table needed?, padding 9
      Next
    Catch ex As Exception
      Email.LogError(ex)
      Console.WriteLine(ex.Message)
      Console.Read()
      Exit Sub
    End Try

    Dim filename As String = String.Empty

    If passheadDatatable.Rows.Count = 1 Then
      filename = String.Format("{0}_mfi-{1}.txt", passheadDatatable.Rows(0).Item("XPHCUST"), passheadDatatable.Rows(0).Item("XPHINV"))
    Else
      filename = String.Format("{0}_mfi.txt", passheadDatatable.Rows(0).Item("XPHCUST"))
    End If
    WriteEDI(filename, ediStringBuilder.ToString(), send_to_ftp)

    Console.Read()
  End Sub

  Private Sub WriteEDI(ByVal filename As String, ByVal ediString As String, ByVal sendAsFTP As Boolean)
    Try
      If sendAsFTP Then
        FtpUploadEDI(filename, ediString.ToString())
        Email.JobCompleteEmail()
      Else
        My.Computer.FileSystem.WriteAllText(filename, ediString.ToString(), False)
        Console.WriteLine("EDI written to text file.")
      End If          '
     
      Console.WriteLine("Complete.")
    Catch ex As Exception
      'Email.LogError(ex)
      Console.WriteLine(ex.Message)
    End Try
  End Sub

  Private Sub FtpUploadEDI(ByVal uploadname As String, ByVal content As String)
    'get a temp file name
    Dim fname As String = System.IO.Path.GetTempFileName
    'write datatable to temp file name in temp location
    Using fs As New IO.StreamWriter(fname)
      fs.Write(content.ToString)
    End Using

    'setup ftp request to upload file with login
    Dim site As String = ftp_mamvast_upload_url_dev & "/" & uploadname
    Dim ftp As FtpWebRequest = FtpWebRequest.Create(site)
    ftp.Credentials = New NetworkCredential(ftp_mamvast_upload_login_dev, ftp_mamvast_upload_pass_dev)
    ftp.Method = WebRequestMethods.Ftp.UploadFile
    ftp.UseBinary = True

    Dim info As New System.IO.FileInfo(fname)
    ftp.ContentLength = info.Length
    Dim buffLength As Integer = 2048
    Dim buff(buffLength - 1) As Byte
    'open the file to be read in a stream
    Using fs As System.IO.FileStream = info.OpenRead()
      'use ftp request stream to create a file for writing
      Using stream As System.IO.Stream = ftp.GetRequestStream()
        Dim contentLen As Integer = fs.Read(buff, 0, buffLength)

        ' Till Stream content ends
        Do While contentLen <> 0
          ' Write Content from the file stream to the FTP Upload Stream
          stream.Write(buff, 0, contentLen)
          contentLen = fs.Read(buff, 0, buffLength)
        Loop
      End Using
    End Using
    'delete temp file in temp location
    System.IO.File.Delete(fname)
  End Sub
End Module
