Imports Microsoft.Office.Interop

Class MainWindow

    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        Dim unit As String
        Dim conn As ADODB.Connection
        Dim path As String
        Dim objExcel As Excel.Application
        Dim demo_mode As Boolean

        path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        conn = New ADODB.Connection
        objExcel = New Excel.Application
        demo_mode = False

        Try
            conn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\2UA3420JXM\Security Group Mapping\WiW Security Group Database Step 2.accdb")
        Catch ex As Exception
            MessageBox.Show("I couldn't open the database.")
        End Try

        If radioButton1.IsChecked = True Then
            unit = "unitID"
            'generate_correction_report(objExcel, conn, UnitSelectionComboBox.SelectedItem.Tag.ToString, unit, path, demo_mode)  'file name, 
            Dim unitid = Trim(Strings.Left(comboBox.SelectedItem.ToString, 3))
            Debug.WriteLine(unitid)
            generate_correction_report(objExcel, conn, unitid, unit, path, demo_mode)  'file name, 
            Generate_Mail_Message(unitid, conn)
            'MessageBox.Show("Unit reports have been saved to" & Chr(10) & path & ".")
            'Else radioButton2.IsChecked = True
            '   unit = "[Change Manager]"
            '  generate_correction_report(objExcel, conn, UnitSelectionComboBox.SelectedItem.Tag.ToString, unit, path, demo_mode)  'file name, 
            ' Generate_Mail_Message(UnitSelectionComboBox.SelectedItem.Tag.ToString, conn)
            'MessageBox.Show("Unit reports have been saved to" & Chr(10) & path & ".")
        Else
            MessageBox.Show("Please select a type of report to generate.")
        End If

    End Sub

    'Private Sub UnitSelectionComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles UnitSelectionComboBox.SelectionChanged
    'MessageBox.Show(UnitSelectionComboBox.SelectedItem.Tag.ToString)
    'End Sub

    Function generate_correction_report(objExcel, conn, where_clause, where_field, folder, demo_mode)
        Dim file_path As String
        Dim rec As ADODB.Recordset
        Dim file_ext As String
        Dim workbook
        Dim worksheet
        Dim file_name_append As String
        Dim sSql As String
        Dim Condition As String
        Dim i As Integer
        Dim j As Integer
        Dim unit As String
        Dim record_count As Integer
        Dim file_name As String
        Dim debug_state As Boolean
        Dim unitID As Integer
        Dim unit_cm As String

        file_path = ""
        rec = New ADODB.Recordset
        file_ext = ".xlsx"
        file_name_append = ""
        sSql = ""
        Condition = ""
        i = 0
        j = 0
        unit = ""
        record_count = 0
        file_name = "WiW Security Group Corrections"
        debug_state = False
        unitID = 0
        unit_cm = ""
        file_name_append = ""


        If where_field = "[Change Manager]" Then
            sSql = "SELECT * FROM Change_manager_summary" & Condition
        Else
            If where_clause = "" Then
                Condition = ""
                sSql = "SELECT count(unitID) AS ""Unit Count"", """ &
                        "All Units"" AS ""Unit"", " &
                        " SUM(EID_Ct) As ""EID_Ct"", " &
                        """All CMs"" As ""Change Manager"" " &
                        "FROM Role_correction_summary" &
                        Condition
            Else
                If where_field = "unitID" Then
                    Condition = " WHERE " & where_field & " = " & where_clause
                ElseIf where_field = "unit"
                    Condition = " WHERE " & where_field & " = """ & where_clause & """ "
                End If
                sSql = "Select * FROM Role_correction_summary" & Condition
            End If

        End If

        Try
            rec.Open(sSql, conn)
        Catch ex As Exception
            MessageBox.Show("Coudn't Find Records.")
        End Try

        j = 0
        If (rec.BOF And rec.EOF) Then
            Debug.WriteLine("No records found.")
        Else
            Do While Not rec.EOF
                i = 0
                If where_field = "[Change Manager]" Then
                    For Each fld In rec.Fields
                        If i = 0 Then
                            unit_cm = fld.value
                            file_name_append = "_" & unit_cm
                        Else
                            record_count = CInt(fld.value)
                        End If
                        i = i + 1
                    Next fld
                Else
                    For Each fld In rec.Fields
                        If i = 0 Then
                            unitID = CInt(fld.value)
                        ElseIf i = 1 Then
                            unit = fld.value
                        ElseIf i = 2 Then
                            record_count = CInt(fld.value)
                            file_name_append = "_" & unit
                        Else
                            unit_cm = fld.value.ToString
                        End If
                        i = i + 1
                    Next fld
                End If

                'i = 0
                j = j + 1
                rec.MoveNext()
            Loop
        End If
        rec.Close()

        file_path = folder & "\" & file_name & file_name_append & file_ext
        file_path = Replace(file_path, "&", "")

        If debug_state = True Then
            objExcel.Visible = True
        End If
        objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth

        Try
            workbook = objExcel.Workbooks.Add
        Catch ex As Exception
            MessageBox.Show("I couldn't add a worksheet!")
        End Try

        Try
            worksheet = workbook.Worksheets("Sheet1")
        Catch ex As Exception
            MessageBox.Show("I couldn't activate the new worksheet!")
        End Try

        Try
            worksheet.Name = "Groups"
        Catch ex As Exception
            MessageBox.Show("I couldn't rename the worksheet!")
        End Try

        Try
            workbook.SaveAs(FileName:=file_path)
        Catch ex As Exception
            MessageBox.Show("File was open.")
        End Try

        generate_by_role_report(objExcel, conn, where_clause, where_field, file_path, unit, "Groups", record_count, demo_mode, workbook)

        If demo_mode = False Then
            workbook.Close()
        End If


        workbook = Nothing
        worksheet = Nothing
        folder = Nothing
        file_ext = Nothing
        file_path = Nothing

    End Function

    Function generate_by_role_report(objExcel, conn, where_clause, where_field, file_path, unit, worksheet_name, record_count, demo_mode, workbook)
        Dim sSql As String
        Dim rec As ADODB.Recordset
        Dim worksheet
        Dim condition As String
        Dim index As Integer
        Dim code As String
        Dim title As String
        Dim i As Integer
        Dim j As Integer
        Dim debug_state As Boolean
        Dim data_column_ct As Integer
        Dim column_offset As Integer
        Dim header_rows As Integer
        Dim role_description As String
        Dim role_array As String()
        Dim foo As String
        Dim formatted_role_description As String
        Dim footer As Boolean

        sSql = ""
        rec = New ADODB.Recordset
        'workbook
        'worksheet
        condition = ""
        index = 0
        code = ""
        title = ""
        i = 0
        j = 0
        debug_state = False
        data_column_ct = 0
        column_offset = 7
        header_rows = 2
        role_description = ""
        foo = ""
        formatted_role_description = ""
        footer = False

        If where_clause = "" Then
            condition = ""
        Else
            If where_field = "unit" Then
                condition = " WHERE " & where_field & " = """ & where_clause & """"
            ElseIf where_field = "unitID" Then
                condition = " WHERE " & where_field & " = " & where_clause
            End If
        End If

        sSql = "Select * FROM Role_correction_report" & condition

        Try
            rec.Open(sSql, conn)
        Catch ex As Exception
            MessageBox.Show("Couldn't Open Role Correction Report")
        End Try

        generate_worksheet(objExcel, rec, file_path, worksheet_name, workbook)
        rec.Close()

        worksheet = workbook.Worksheets(worksheet_name)

        If debug_state = True Then
            objExcel.Visible = True
            worksheet.Activate
        End If

        If demo_mode = True Then
            objExcel.Visible = True
            worksheet.Activate
        End If

        worksheet.Rows("1").Insert

        sSql = "Select role_code, role_title, role_description FROM roles WHERE role_order Is Not null ORDER BY  `role_order` asc"
        'Debug.WriteLine(sSql)
        rec.Open(sSql, conn)
        If (rec.BOF And rec.EOF) Then
            Debug.WriteLine("No records found.")
        Else
            Do While Not rec.EOF
                i = 0
                formatted_role_description = ""
                For Each fld In rec.Fields
                    If i = 0 Then
                        code = fld.value
                    ElseIf i = 1 Then
                        title = fld.value
                    Else
                        role_description = fld.value
                        role_array = Split(role_description, "*")
                        Dim foos = role_array.Count
                        foos = foos - 1
                        Dim ii = 0
                        For Each foo In role_array
                            If ii < foos Then
                                formatted_role_description = formatted_role_description & Trim(foo) & Chr(10) & "   - "
                            Else
                                formatted_role_description = formatted_role_description & Trim(foo)
                            End If
                            ii = ii + 1
                        Next foo
                    End If
                    i = i + 1
                Next fld
                i = 0
                j = j + 1
                worksheet.cells(1, j + column_offset).Value = code
                worksheet.cells(2, j + column_offset).Value = title
                If footer = True Then
                    worksheet.cells(header_rows + record_count + 1, j + column_offset).Value = formatted_role_description
                End If
                rec.MoveNext()
            Loop
        End If
        rec.Close()
        data_column_ct = j

        'Range Definitions
        Dim max_column = column_offset + data_column_ct
        Dim max_row = header_rows + record_count
        Dim max_row_address = worksheet.Rows(max_row).Address
        Dim max_column_txt = worksheet.Cells(1, max_column).Address
        Dim max_cell_txt = worksheet.Cells(max_row, max_column).Address
        Dim max_header_txt = worksheet.Cells(header_rows, max_column).Address
        Dim data_header_start = worksheet.Cells(1, column_offset).Address
        Dim data_columns = column_offset + 1 & ":" & data_column_ct
        Dim Dataset = worksheet.Range("A" & header_rows + 1 & ": " & max_cell_txt)
        Dim entire_sheet = worksheet.Range("A1:" & max_cell_txt)
        Dim footer_row = max_row + 1
        Dim Data_columns_address = worksheet.Range(worksheet.Columns(column_offset + 1), worksheet.Columns(max_column)).Address

        'Column Offset Modifications
        'UnitID
        With worksheet.Columns("A:A")
            .ColumnWidth = 3
        End With
        'Unit
        With worksheet.Columns("B:B")
            .ColumnWidth = 38
            .EntireColumn.Hidden = True
        End With
        'Change Manager
        With worksheet.Columns("C:C")
            .ColumnWidth = 8
            .EntireColumn.Hidden = True
        End With
        'Budget Number
        With worksheet.Columns("D:D")
            .ColumnWidth = 15
            .WrapText = True
        End With
        'EID
        With worksheet.Columns("E:E")
            .ColumnWidth = 10
        End With
        'Employee Name
        With worksheet.Columns("F:F")
            .ColumnWidth = 25
            .WrapText = True
        End With
        'Supervisory Org
        With worksheet.Columns("G:G")
            .ColumnWidth = 40
            .WrapText = True
        End With
        'Data Columns
        If footer = True Then
            With worksheet.Columns(Data_columns_address)
                .ColumnWidth = 40
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With
            With worksheet.Rows(footer_row)
                .Font.Size = 8
                .Font.ColorIndex = 16
                .WrapText = True
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                .VerticalAlignment = Excel.XlVAlign.xlVAlignTop
            End With
        Else
            With worksheet.Columns(Data_columns_address)
                .ColumnWidth = 4
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With
        End If

        'Dataset Color Coding
        index = column_offset
        Do
            If worksheet.Cells(1, index).Value = "I9" Then
                worksheet.Columns(index).Interior.Color = RGB(253, 228, 207)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(header_rows, index)).Interior.Color = RGB(247, 150, 70)
            ElseIf worksheet.Cells(1, index).Value = "ABP" Then
                worksheet.Columns(index).Interior.Color = RGB(218, 231, 246)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(header_rows, index)).Interior.Color = RGB(83, 141, 213)
            ElseIf worksheet.Cells(1, index).Value = "ACP" Then
                worksheet.Columns(index).Interior.Color = RGB(246, 230, 230)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(header_rows, index)).Interior.Color = RGB(218, 150, 148)
            ElseIf worksheet.Cells(1, index).Value = "CP" Then
                worksheet.Columns(index).Interior.Color = RGB(238, 234, 242)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(header_rows, index)).Interior.Color = RGB(128, 100, 162)
            ElseIf worksheet.Cells(1, index).Value = "CAC" Then
                worksheet.Columns(index).Interior.Color = RGB(228, 223, 236)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(header_rows, index)).Interior.Color = RGB(228, 223, 236)
            ElseIf worksheet.Cells(1, index).Value = "HRC" Then
                worksheet.Columns(index).Interior.Color = RGB(228, 228, 228)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(header_rows, index)).Interior.Color = RGB(178, 178, 178)
            ElseIf worksheet.Cells(1, index).Value = "HRP" Then
                worksheet.Columns(index).Interior.Color = RGB(205, 233, 239)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(header_rows, index)).Interior.Color = RGB(49, 134, 155)
            ElseIf worksheet.Cells(1, index).Value = "TC" Then
                worksheet.Columns(index).Interior.Color = RGB(241, 245, 231)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(header_rows, index)).Interior.Color = RGB(196, 215, 155)
            End If
            index = index + 1
        Loop Until index > max_column

        'Header Modifications
        worksheet.Range("A1: " & max_header_txt).Font.Bold = True
        worksheet.Range("A2:" & max_header_txt).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
        worksheet.Range("H2:" & max_header_txt).Orientation = 90

        'All Cells
        With worksheet.Range("A1:" & max_cell_txt).Font
            .Size = 10
        End With

        'All Data Rows
        worksheet.Range("A" & header_rows + 1 & ": " & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlDot
        worksheet.Range("A" & header_rows + 1 & ": " & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).ThemeColor = 1
        worksheet.Range("A" & header_rows + 1 & ": " & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = -0.14996795556505
        worksheet.Rows("3:" & max_row).Autofit

        'AutoFilter
        worksheet.Range("A2:" & max_cell_txt).Autofilter

        'Page Setup
        worksheet.PageSetup.PrintArea = "$A$1:" & max_cell_txt
        worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaper11x17
        worksheet.PageSetup.PrintTitleRows = "$1:$2"
        worksheet.PageSetup.PrintTitleColumns = "$A:$G"
        worksheet.PageSetup.CenterHeader = unit & Chr(10) & "Correction Proof of " & worksheet_name
        worksheet.PageSetup.RightHeader = "&D"
        worksheet.PageSetup.LeftFooter = "This worksheets presents changes identified by the unit, subunit or department Readiness Lead in coordination with the unit's Change Manager."
        worksheet.PageSetup.RightFooter = "&P of &N"

        workbook.Save()

        If demo_mode = True Then
            Threading.Thread.CurrentThread.Sleep(500)
        End If

        sSql = Nothing
        rec = Nothing
        workbook = Nothing
        worksheet = Nothing

    End Function

    Function generate_worksheet(objExcel, recordset, file_path, worksheet_name, workbook)
        Dim Worksheet
        Dim fieldCount
        Dim recArray
        Dim recCount
        Dim debug_state

        debug_state = False

        If debug_state = True Then
            objExcel.Visible = True
        End If

        objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth
        Worksheet = workbook.Worksheets(worksheet_name)

        ' Copy field names to the first row of the worksheet
        fieldCount = recordset.Fields.Count
        For iCol = 1 To fieldCount
            Worksheet.Cells(1, iCol).Value = recordset.Fields(iCol - 1).Name
        Next

        ' Check version of Excel
        If Val(Mid(objExcel.Version, 1, InStr(1, objExcel.Version, ".") - 1)) > 8 Then
            'EXCEL 2000,2002,2003, or 2007: Use CopyFromRecordset

            ' Copy the recordset to the worksheet, starting in cell A2
            Worksheet.Cells(2, 1).CopyFromRecordset(recordset)
            'Note: CopyFromRecordset will fail if the recordset
            'contains an OLE object field or array data such
            'as hierarchical recordsets

        Else
            'EXCEL 97 or earlier: Use GetRows then copy array to Excel

            ' Copy recordset to an array
            recArray = recordset.GetRows
            'Note: GetRows returns a 0-based array where the first
            'dimension contains fields and the second dimension
            'contains records. We will transpose this array so that
            'the first dimension contains records, allowing the
            'data to appears properly when copied to Excel

            ' Determine number of records

            recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array


            ' Check the array for contents that are not valid when
            ' copying the array to an Excel worksheet
            For iCol = 0 To fieldCount - 1
                For iRow = 0 To recCount - 1
                    ' Take care of Date fields
                    If IsDate(recArray(iCol, iRow)) Then
                        recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                        ' Take care of OLE object fields or array fields
                    ElseIf IsArray(recArray(iCol, iRow)) Then
                        recArray(iCol, iRow) = "Array Field"
                    End If
                Next iRow 'next record
            Next iCol 'next field

            ' Transpose and Copy the array to the worksheet,
            ' starting in cell A2
            Worksheet.Cells(2, 1).Resize(recCount, fieldCount).Value =
                    TransposeDim(recArray)
        End If

        ' Auto-fit the column widths and row heights
        'objExcel.Selection.CurrentRegion.Columns.AutoFit
        objExcel.Selection.CurrentRegion.Rows.AutoFit

        workbook.SaveAs(FileName:=file_path)

        workbook = Nothing
        Worksheet = Nothing

    End Function

    Function TransposeDim(v As Object) As Object
        ' Custom Function to Transpose a 0-based array (v)

        Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
        Dim tempArray As Object

        Xupper = UBound(v, 2)
        Yupper = UBound(v, 1)

        ReDim tempArray(Xupper, Yupper)
        For X = 0 To Xupper
            For Y = 0 To Yupper
                tempArray(X, Y) = v(Y, X)
            Next Y
        Next X

        TransposeDim = tempArray
    End Function

    Function Generate_Mail_Message(unitID, conn)
        Dim appOutlook As Outlook.Application
        Dim msg As Outlook.MailItem
        Dim rec As ADODB.Recordset
        Dim sSql As String
        Dim unit_cm As String
        Dim sub_unit As String
        Dim unit As String
        Dim first As String
        Dim email As String

        unit = ""
        sub_unit = ""
        unit_cm = "Copp"

        rec = New ADODB.Recordset
        appOutlook = New Outlook.Application

        sSql = "Select unit_name, sub_unit, unit_cm from units where unitid = " & unitID
        rec.Open(sSql, conn)
        Dim i = 0
        If (rec.BOF And rec.EOF) Then
            Debug.WriteLine("No records found.")
        Else
            For Each fld In rec.Fields
                If i = 0 Then
                    unit = fld.value
                ElseIf i = 1 Then
                    sub_unit = fld.value
                ElseIf i = 2 Then
                    unit_cm = fld.value.ToString
                End If
                i = i + 1
            Next fld
        End If

        If unit_cm = "Copp" Then
            email = "coppk@uw.edu"
            first = "Karen"
        ElseIf unit_cm = "Wiggers" Then
            email = "hw15@uw.edu"
            first = "Hannah"
        ElseIf unit_cm = "Toledo" Then
            email = "vToledo@uw.edu"
            first = "Valerie"
        ElseIf unit_cm = "Mow" Then
            email = "stevemow@uw.edu"
            first = "Steven"
        ElseIf unit_cm = "Greemwood" Then
            email = "gail1@uw.edu"
            first = "Gail"
        End If

        Dim subject = "Role Mapping Corrections for " & sub_unit & " are ready to go"

        Dim msg_body = "Hello " & first & ",<br><br>" &
                        "I wanted to let you know that Sophie and I have wrapped up incorporating the changes to the " & sub_unit & " dataset that you and your Readiness Lead identified.<br><br>" &
                        "The document has been saved to the old SharePoint site, and is available here: " &
                        "<a href=""https://sharepoint.washington.edu/oim/proj/HRPayroll/Imp/Supervisory%20Org%20Cleanup/Role_mapping_2/Corrections/WiW%20Security%20Group%20Corrections_" & Replace(sub_unit, " ", "%20") & "%20(" & Replace(unit, " ", "%20") & ").xlsx"">" &
                        "https//sharepoint.washington.edu/oim/proj/HRPayroll/Imp/Supervisory%20Org%20Cleanup/Role_mapping_2/Corrections/WiW%20Security%20Group%20Corrections_" & sub_unit & "%20(" & unit & ").xlsx</a><br><br>" &
                        "I'll leave a hard copy of the changes on your deskchair.<br><br>" &
                        "Please let me know if I can be of further assistance.<br><br>" &
                        "Bill Bascus"

        msg = appOutlook.CreateItem(Outlook.OlItemType.olMailItem)
        msg.To = email
        msg.CC = "slsp@uw.edu"
        msg.Subject = subject
        msg.HTMLBody = msg_body
        msg.Display(True)

    End Function

    Sub Generate_UnitSelectionComboBox2()
        Dim rec As ADODB.Recordset
        Dim conn As ADODB.Connection
        Dim sSql As String
        Dim cbox As System.Windows.Controls.ComboBox
        Dim cboxitem As System.Windows.Controls.ComboBoxItem
        Dim tag As String
        Dim content As String


        cbox = New ComboBox

        rec = New ADODB.Recordset
        conn = New ADODB.Connection

        conn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\2UA3420JXM\Security Group Mapping\WiW Security Group Database Step 2.accdb")

        sSql = "SELECT units.unitID, units.unit_name  & "": "" & units.sub_unit FROM units " &
                    "WHERE sub_unit <> ""A-All Roll-Up"" " &
                    "And Ready_For_WD = True " &
                    "ORDER BY units.unit_name, units.sub_unit"

        Try
            rec.Open(sSql, conn)
        Catch
            Debug.WriteLine(sSql)
        End Try

        If (rec.BOF And rec.EOF) Then
            Debug.WriteLine("No records found.")
        Else
            Do While Not rec.EOF
                Dim i = 0
                For Each fld In rec.Fields
                    If i = 0 Then
                        tag = fld.value.ToString
                        'Debug.WriteLine(tag)
                    Else
                        content = Replace(fld.value.ToString, "&", "")
                        content = Replace(content, "-", "")
                        'Debug.WriteLine(content)
                    End If


                    i = i + 1
                Next fld
                cboxitem = New ComboBoxItem()
                cboxitem.Tag = tag
                cboxitem.Content = content
                cboxitem.HorizontalContentAlignment = HorizontalAlignment.Left
                cboxitem.VerticalContentAlignment = VerticalAlignment.Top

                Try
                    comboBox.Items.Add(tag & " - " & content)
                    'cbox.Items.Add(cboxitem)
                Catch ex As Exception
                    Debug.WriteLine(ex.ToString)
                End Try
                rec.MoveNext()
            Loop
        End If

        Try
            'Debug.WriteLine(canvas1.Children.Count.ToString)
            'Me.canvas1.Children.Add(cbox)
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try

    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Generate_UnitSelectionComboBox2()
    End Sub

End Class
