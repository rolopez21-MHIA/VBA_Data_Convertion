Attribute VB_Name = "Module1"
Option Explicit

Dim wp1 As Worksheet
Dim wp2 As Worksheet
Dim wp3 As Worksheet
Dim Lrow As Double
Dim Lrow2 As Double
Dim Lrow3 As Double
Dim LCol As Double
Dim LCol2 As Double
Dim x As Variant
Dim TrgVal As String

Sub Data_Hub()

Application.ScreenUpdating = False
        
        Set wp2 = Sheets("Income")
        Set wp3 = Sheets("wp")
        Set wp1 = Sheets("Outcome")

        wp3.Range("A:ZZ").Clear
        wp1.Range("A:ZZ").Clear
        
            LCol = wp2.Cells(2, Columns.Count).End(xlToLeft).Column
            LCol2 = wp2.Cells(1, Columns.Count).End(xlToLeft).Column
            Lrow = wp2.Cells(Rows.Count, 1).End(xlUp).Row
            
            wp2.Range(wp2.Cells(2, "A"), wp2.Cells(2, LCol2)).Copy wp1.Range("A1")
            
            
                For x = LCol To LCol2 + 1 Step -1
                
                    TrgVal = wp2.Cells(2, x)
                    
                    Lrow2 = wp3.Cells(Rows.Count, 1).End(xlUp).Row + 1
                    Lrow3 = wp3.Cells(Rows.Count, LCol2 + 1).End(xlUp).Row + 1
                
                            wp2.Range(wp2.Cells(3, "A"), wp2.Cells(Lrow, LCol2)).Copy
                            wp3.Cells(Lrow2, 1).PasteSpecial xlPasteValues
                            
                    Lrow2 = wp3.Cells(Rows.Count, 1).End(xlUp).Row
                            
                            wp3.Range(wp3.Cells(Lrow3, LCol2 + 1), wp3.Cells(Lrow2, LCol2 + 1)).Value = TrgVal
                            
                            wp2.Range(wp2.Cells(3, x), wp2.Cells(Lrow, x)).Copy
                            wp3.Range(wp3.Cells(Lrow3, LCol2 + 2), wp3.Cells(Lrow2, LCol2 + 2)).PasteSpecial xlPasteValues
                    
                
                Next
                
            Lrow3 = wp3.Cells(Rows.Count, 1).End(xlUp).Row
            'LCol = wp3.Cells(2, Columns.Count).End(xlToLeft).Column
                
                
                For x = Lrow3 To 2 Step -1
                
                    If wp3.Cells(x, LCol2 + 2) = "" Or wp3.Cells(x, LCol2 + 2) = 0 Then
                    
                        wp3.Cells(x, LCol).EntireRow.Delete
                    
                    End If
                
                Next
            
            Lrow3 = wp3.Cells(Rows.Count, 1).End(xlUp).Row
            Lrow2 = wp1.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
                wp3.Range(wp3.Cells(2, "A"), wp3.Cells(Lrow3, LCol2 + 2)).Copy
                wp1.Cells(Lrow2, 1).PasteSpecial xlPasteValues
                
        
 Application.ScreenUpdating = True
                

End Sub



