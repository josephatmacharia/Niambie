NIAMBIE V1.1

Is an Excel Vba application that helps Accountants 
reconcile large supplier or customer data sets. 

Sub reconcile()
      
      'Disable Screen Updating and Events
       ThisWorkbook.Application.EnableEvents = False
       ThisWorkbook.Application.ScreenUpdating = False
       
      'Display the progress bar
       IncrementalProgress.Show
       
      'Prevent the mouse from flickering
       ThisWorkbook.Application.Cursor = xlNorthwestArrow
      
      'Calls procedure to match refrences and amounts
       Call match_references_and_amounts
      
      'Calls procedure to sort columns
       Call sort_columns
      
       'Call procedure to move amounts with differences
       Call move_amounts_with_differences
     
      'Calls the procedure and computes the differences
       Call calculate_differences_in_amount
    
       'Call method to take values to SELF CHECK
       Call calculate_totals_and_move_torecon
        
       'Stop the progress bar
       Unload IncrementalProgress
       
      'Display message to show completion
       MsgBox "Imemaliza fiti ", vbOKOnly, "Niambie"
    
      'Disable the reconcile,reconcile and ledger button
       With Worksheets("MENU")
           .cmd_reconcile.Enabled = False
           .cmd_ledger.Enabled = False
           .cmd_stmnt.Enabled = False
           .cmd_refresh.Enabled = True
           .cmd_reconciliation.Enabled = True
       End With
      
      'Enable Screen Updating and Events
       ThisWorkbook.Application.EnableEvents = True
       ThisWorkbook.Application.ScreenUpdating = True
    
End Sub

'Procedure to match invoice numbers and amounts
Sub match_references_and_amounts()
 
       Dim i, j, k, lrow_our_ledger, lrow_sup_invoices_reconciled, amount, inv_no, lrow_supplier_stmnt, lrow_pricediff_ourside, lrow_pricediff_supp_side, lrow_reconciled As Long
       Dim sPercentage As Single
       Dim sStatus As String
       
        'Find the last row that contains data.in column A
       With Worksheets("our_ledger")
          lrow_our_ledger = .Cells(.Rows.Count, "A").End(xlUp).Row
       End With
        
       'Find the last row that contains data.in column A
       With Worksheets("supplier_stmt")
          lrow_supplier_stmnt = .Cells(.Rows.Count, "A").End(xlUp).Row
       End With
        
        'first row number where you need to paste values'
        With Worksheets("reconciled_invoices_our_side")
            lrow_reconciled = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
        
         'first row number where you need to paste  values'
         With Worksheets("reconciled_supplier_side")
            lrow_sup_invoices_reconciled = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
        
        'first row number where you need to paste  values'
        With Worksheets("pricediff_our_side")
            lrow_pricediff_ourside = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
        
        
         'first row number where you need to paste  values'
         With Worksheets("pricediff_supp_side")
            lrow_pricediff_supp_side = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
        
       'cycle through every row in our ledger upto to the last cell with data
        For i = 2 To lrow_our_ledger
                
             'Grab values from the reference/inv_no column and amount
              inv_no = Worksheets("our_ledger").Cells(i, 1)
              amount = Worksheets("our_ledger").Cells(i, 4)
              
              'cycle through every row in our supplier_stmnt upto to the last cell with data
              For k = 2 To lrow_supplier_stmnt
                
                'Check if the inv_number and amount matches
                If Worksheets("supplier_stmt").Cells(k, 1) = inv_no And Worksheets("supplier_stmt").Cells(k, 4) = amount Then
                    
                    'Cut and paste matched values to respective worksheets
                     Worksheets("our_ledger").Rows(i).Cut Destination:=Worksheets("reconciled_invoices_our_side").Range("A" & lrow_reconciled)
                     lrow_reconciled = lrow_reconciled + 1
                    
                      Worksheets("supplier_stmt").Rows(k).Cut Destination:=Worksheets("reconciled_supplier_side").Range("A" & lrow_sup_invoices_reconciled)
                      lrow_sup_invoices_reconciled = lrow_sup_invoices_reconciled + 1
                      Exit For
                 'If invoice matches and amount does not cut and paste the data to respective sheets
                  ElseIf Worksheets("supplier_stmt").Cells(k, 1) = inv_no And Worksheets("supplier_stmt").Cells(k, 4) <> amount Then
    
                      Worksheets("our_ledger").Rows(i).Cut Destination:=Worksheets("pricediff_our_side").Range("A" & lrow_pricediff_ourside)
                      lrow_pricediff_ourside = lrow_pricediff_ourside + 1
    
                      Worksheets("supplier_stmt").Rows(k).Cut Destination:=Worksheets("pricediff_supp_side").Range("A" & lrow_pricediff_supp_side)
                      lrow_pricediff_supp_side = lrow_pricediff_supp_side + 1
                      Exit For
                  
                
                End If
              
              Next k
              
              'calculate the number of steps and pass values to Increment function
              sPercentage = Round((i / lrow_our_ledger) * 100)
              sStatus = "Chill kiasi!....Processing: " & i & " of " & lrow_our_ledger & " records: "
             
             'Must DoEvents to allow code to update bar and show it
             DoEvents
              
             'Calls the progress bar for each task completed
            IncrementalProgress.Increment sPercentage, sStatus
             
              'Application.StatusBar = "Processing: " & i & "of" & lrow_our_ledger & "records: " & Format(i / lrow_our_ledger, "0%")
             
         Next i
            'Application.StatusBar = False
End Sub

'Sort columns
Sub sort_columns()
     Dim i, j, lrow_our_ledger, lrow_supplier_stmnt, lrow_reconciled As Long
     
      'Find the last row that contains data.in column A
     With Worksheets("our_ledger")
        lrow_our_ledger = .Cells(.Rows.Count, "A").End(xlUp).Row
     End With
      
     'Find the last row that contains data.in column A
     With Worksheets("supplier_stmt")
        lrow_supplier_stmnt = .Cells(.Rows.Count, "A").End(xlUp).Row
     End With
     
     'sort both our_ledger and supplier statement and transfer values to unreceived and unclaimed ledgers respectively
     Sheets("our_ledger").Range("A2:D" & lrow_our_ledger).Sort key1:=Sheets("our_ledger").Range("A2:A" & lrow_our_ledger), order1:=xlAscending, Header:=xlNo
     Sheets("our_ledger").Range("A2:D" & lrow_our_ledger).Cut Destination:=Worksheets("unclaimed").Range("A2")
    
     Sheets("supplier_stmt").Range("A2:D" & lrow_supplier_stmnt).Sort key1:=Sheets("supplier_stmt").Range("A2:A" & lrow_supplier_stmnt), order1:=xlAscending, Header:=xlNo
     Sheets("supplier_stmt").Range("A2:D" & lrow_supplier_stmnt).Cut Destination:=Worksheets("unreceived").Range("A2")
    
End Sub

Sub move_amounts_with_differences()
    'first row number where you need to paste  values'
        With Worksheets("pricediff_our_side")
            lrow_pricediff_ourside = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
        
         'first row number where you need to paste  values'
         With Worksheets("pricediff_supp_side")
            lrow_pricediff_supp_side = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
     
     'copy price differences from one worksheet to another
     Sheets("pricediff_supp_side").Range("A2:E" & lrow_pricediff_supp_side).Cut Destination:=Worksheets("price_differences").Range("F4")
     Sheets("pricediff_our_side").Range("A2:E" & lrow_pricediff_ourside).Cut Destination:=Worksheets("price_differences").Range("A4")

End Sub


'Procedure to compute differences in amount
Sub calculate_differences_in_amount()

    Dim i As Integer
    Dim lrow_in_pricediff  As Long
    
     With Worksheets("price_differences")
      'Find the last row that contains data.in column A
        lrow_in_pricediff = .Cells(.Rows.Count, "A").End(xlUp).Row
        'Cycle through every column upto the last row
         For i = 4 To lrow_in_pricediff
           'Compute the difference in amounts and store the value in column K
           .Range("K" & i).Value = .Range("I" & i).Value - .Range("D" & i).Value
            
            'Check for negatives
            If .Cells(i, 11).Value < 0 Then
              'Store as undercharge and set font color to none
              
                .Cells(i, 12).Value = "undercharge"
                .Cells(i, 12).Font.ColorIndex = 0
             
            Else
            
             'Store as overcharge and set font color to red
             .Cells(i, 12).Value = "overcharge"
             .Cells(i, 12).Font.ColorIndex = 3
             
            End If
            
         Next i
         
       'Format column 11 as accounting
        .Columns(11).NumberFormat = "#,##0.00"
     End With

End Sub


Sub calculate_totals_and_move_torecon()
   'loop through unreceived and calculate total
    Dim total_unreceived, total_unclaimed, total_overcharge, total_undercharge, _
     total_remitance, total_supp_stmnt, total_pricediff_ourledger, total_pricediff_supplier As Long
   
   'Find the last row that contains data.in column A
    With Worksheets("unreceived")
       lrow_unreceived = .Cells(.Rows.Count, "D").End(xlUp).Row
    End With
    'calculate the total of unreceived invoices
    total_unreceived = Application.WorksheetFunction.Sum(Sheets("unreceived").Range("D2:D" & lrow_unreceived))
    'take the total value computed to cell g17
    Sheets("RECON").Range("G17").Value = total_unreceived
    
   'Find the last row that contains data.in column A
    With Worksheets("unclaimed")
       lrow_unclaimed = .Cells(.Rows.Count, "D").End(xlUp).Row
    End With
    
    'calculate the total of unclaimed invoices
    total_unclaimed = Application.WorksheetFunction.Sum(Sheets("unclaimed").Range("D2:D" & lrow_unclaimed))
    'take the total value computed to cell n17
    Sheets("RECON").Range("N17").Value = total_unclaimed
    
    'Find the last row that contains data.in column A
    With Worksheets("price_differences")
       lrow_pricediff = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
     
     
     'compute the total of overcharged invoices only
     total_overcharge = Application.WorksheetFunction.SumIf(Sheets("price_differences").Range("L4:L" & lrow_pricediff), _
     "overcharge", Sheets("price_differences").Range("K4:K" & lrow_pricediff))
     'take the value computed above to cell g18
     Sheets("RECON").Range("G18").Value = total_overcharge
   
    'compute the total of undercharged invoices only
    total_undercharge = Application.WorksheetFunction.SumIf(Sheets("price_differences").Range("L4:L" & lrow_pricediff), _
    "undercharge", Sheets("price_differences").Range("K4:K" & lrow_pricediff))
   'take the value computed above to cell g12
    Sheets("RECON").Range("G12").Value = total_undercharge
    
   'compute the total of invoices with price differences on ourledger
   total_pricediff_ourledger = Application.WorksheetFunction.Sum(Sheets("price_differences").Range("D4:D" & lrow_pricediff))
   'compute the total of invoices with price differences on supplier side
   total_pricediff_supplier = Application.WorksheetFunction.Sum(Sheets("price_differences").Range("I4:I" & lrow_pricediff))
   
 'Find the last row that contains data.in column A
   With Worksheets("reconciled_invoices_our_side")
       lrow_recon_ourside = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
   'find the total of amount in the worksheet
   total_remitance = Application.WorksheetFunction.Sum(Sheets("reconciled_invoices_our_side").Range("D2:D" & lrow_recon_ourside))
   
   
   'Update balance per supplier side cell O8
    Sheets("RECON").Range("O8").Value = total_remitance + total_unclaimed + total_pricediff_ourledger
    Sheets("RECON").Range("H8").Value = total_remitance + total_unreceived + total_pricediff_supplier
    
    
    
End Sub


Sub refresh()
  'Sub to delete data on the worksheet when the workbook opens
    Dim sht As Worksheet
    'cycle through every worksheet in this workbook
    For Each sht In ThisWorkbook.Sheets
        'Check if the worksheet name is RECON
        If sht.Name <> "RECON" Then
          'Clear all content if the worksheet is not RECON
          sht.UsedRange.ClearContents
        End If
        
    Next sht
    
    'Enable the reconcile,reconcile and ledger button
     With Worksheets("MENU")
         '.cmd_reconcile.Enabled = True
         .cmd_ledger.Enabled = True
         .cmd_stmnt.Enabled = True
     End With

End Sub







