' Honors Inventory Project Part 1
' Jonathan Rogers
' 7/27/2020

' The purpose of this program is to organize product information entered by the user into 
' three or fewer sections, one for each product (this program can only handle three products 
' currently). This program collects information such as product name, product number, the 
' amount in stock, and the price per item. Using this information, the program both organizes 
' the products, but also calculates the total value of each product (in-stock only) and the 
' number of products being displayed (max 3).

' This program uses twelve labels, four text boxes, one button, and one strip menu on a Windows Form.

Public Class Form1


    ' The following vars change based on user interaction
    Dim decTotalValueInStock1 As Decimal ' Stores the product of decPricePerItem1 and intAmountInStock1 for the first product.
    Dim decTotalValueInStock2 As Decimal ' Stores the product of decPricePerItem2 and intAmountInStock2 for the second product.
    Dim decTotalValueInStock3 As Decimal ' Stores the product of decPricePerItem3 and intAmountInStock3 for the third product.

    Dim strProductName1 As String ' Stores product name for the first product.
    Dim strProductName2 As String ' Stores product name for the second product.
    Dim strProductName3 As String ' Stores product name for the third product.

    Dim intProductNumber1 As Integer ' Stores the product number for the first product.
    Dim intProductNumber2 As Integer ' Stores the product number for the second product.
    Dim intProductNumber3 As Integer ' Stores the product number for the third product.

    Dim intAmountInStock1 As Integer ' Stores the amount in stock value for the first product (how many items are in stock).
    Dim intAmountInStock2 As Integer ' Stores the amount in stock value for the second product (how many items are in stock).
    Dim intAmountInStock3 As Integer ' Stores the amount in stock value for the third product (how many items are in stock).

    Dim decPricePerItem1 As Decimal ' Stores the price per item value for the first value.
    Dim decPricePerItem2 As Decimal ' Stores the price per item value for the second value.
    Dim decPricePerItem3 As Decimal ' Stores the price per item value for the third value.


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

        ' Local Variables
        Static intButtonPresses As Integer ' Stores the amount of successful button clicks

        Static intCheck1 As Integer = 0 ' Stores the amount of successful operations completed by computer for product 1. 
        Static intCheck2 As Integer = 0 ' Stores the amount of successful operations completed by computer for product 2.
        Static intCheck3 As Integer = 0 ' Stores the amount of successful operations completed by computer for product 3.

        Select Case intButtonPresses
            Case = 0
                strProductName1 = Me.txtProductName.Text ' Sets the values in txtProductName to a variable
                If strProductName1 = Nothing Then ' If there is no values...
                    MessageBox.Show("Enter a product name") ' ... a message box will display saying "Enter a product name"
                Else ' If there is text...
                    intCheck1 += 1 ' intCheck1 will increase by one symbolizing a successful operation.

                    Try ' Checks if...
                        intProductNumber1 = Me.txtProductNumber.Text ' ...this is a numeric value. If not...
                        intCheck1 += 1 ' ... this will not run (check using a breakpoint)...
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid product number") '... and a message box will show
                    End Try

                    Try ' Checks if...
                        intAmountInStock1 = Me.txtAmountInStock.Text ' ...this is a numeric value. If not...
                        intCheck1 += 1 ' ... this will not run (check using a breakpoint)...
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid inventory amount") '... and a message box will show
                    End Try

                    Try ' Checks if...
                        decPricePerItem1 = Me.txtPricePerItem.Text ' ...this is a numeric value. If not...
                        intCheck1 += 1 ' ... this will not run (check using a breakpoint)...
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid price") '... and a message box will show
                    End Try


                    decTotalValueInStock1 = decPricePerItem1 * intAmountInStock1 ' Calculates total value of product

                    ' This ensures that the program successfully completed the four statements above...
                    If intCheck1 = 4 Then ' If all operations are completed correctly
                        intButtonPresses += 1 ' and adds 1 to the counter. Max amount of button clicks is three.
                        Me.lblProduct1.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                    Else
                        intCheck1 = 0 ' If not all operations were completed successfully, then intCheck1 will be set back to 1.
                    End If
                End If

            Case = 1
                strProductName2 = Me.txtProductName.Text ' Sets the values in txtProductName to a variable
                If strProductName2 = Nothing Then ' If there is no values...
                    MessageBox.Show("Enter a product name") ' ... a message box will display saying "Enter a product name"
                Else ' If there is text...
                    intCheck2 += 1 ' intCheck2 will increase by one symbolizing a successful operation.

                    Try ' Checks if...
                        intProductNumber2 = Me.txtProductNumber.Text ' ...this is a numeric value. If not...
                        If intProductNumber2 = intProductNumber1 Then ' If the same product number is used more than once...
                            MessageBox.Show("Please use a different product number")
                        Else
                            intCheck2 += 1 ' ... this will not run (check using a breakpoint)...
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid product number") '... and a message box will show
                    End Try

                    Try ' Checks if...
                        intAmountInStock2 = Me.txtAmountInStock.Text ' ...this is a numeric value. If not...
                        intCheck2 += 1 ' ... this will not run (check using a breakpoint)...
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid inventory amount") '... and a message box will show
                    End Try

                    Try ' Checks if...
                        decPricePerItem2 = Me.txtPricePerItem.Text ' ...this is a numeric value. If not...
                        intCheck2 += 1 ' ... this will not run (check using a breakpoint)...
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid price") '... and a message box will show
                    End Try


                    decTotalValueInStock2 = decPricePerItem2 * intAmountInStock2 ' Calculates total value of product

                    ' This ensures that the program successfully completed the four statements above...
                    If intCheck2 = 4 Then
                        intButtonPresses += 1 ' and adds 1 to the counter. Max amount of button clicks is three.
                        If intProductNumber1 > intProductNumber2 Then ' If product number 1 is less then product number 2
                            lblProduct1.Text = Nothing ' Clears text in label
                            lblProduct2.Text = Nothing ' Clears text in label
                            Me.lblProduct1.Text = intProductNumber2 & vbCrLf & strProductName2 & vbCrLf & intAmountInStock2 & vbCrLf & "$" & decPricePerItem2 & vbCrLf & "$" & decTotalValueInStock2 ' Outputs the data entered by the user from the left side of the program
                            Me.lblProduct2.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                        ElseIf intProductNumber1 < intProductNumber2 Then ' If product number 1 is less then product number 2
                            lblProduct1.Text = Nothing ' Clears text in label
                            lblProduct2.Text = Nothing ' Clears text in label
                            Me.lblProduct2.Text = intProductNumber2 & vbCrLf & strProductName2 & vbCrLf & intAmountInStock2 & vbCrLf & "$" & decPricePerItem2 & vbCrLf & "$" & decTotalValueInStock2 ' Outputs the data entered by the user from the left side of the program
                            Me.lblProduct1.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                        End If
                    Else
                        intCheck2 = 0 ' If not all operations were completed successfully, then intCheck2 will be set back to 1.
                    End If
                End If
            Case 2
                strProductName3 = Me.txtProductName.Text ' Sets the values in txtProductName to a variable
                If strProductName3 = Nothing Then ' If there is no values...
                    MessageBox.Show("Enter a product name") ' ... a message box will display saying "Enter a product name"
                Else ' If there is text...
                    intCheck3 += 1 ' intCheck3 will increase by one symbolizing a successful operation.

                    Try ' Checks if...
                        intProductNumber3 = Me.txtProductNumber.Text ' ...this is a numeric value. If not...
                        If intProductNumber3 = intProductNumber2 Or intProductNumber3 = intProductNumber1 Then ' If the same product number is used more than once...
                            MessageBox.Show("Please use a different product number")
                        Else
                            intCheck3 += 1 ' ... this will not run (check using a breakpoint)...
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid product number") '... and a message box will show
                    End Try

                    Try ' Checks if...
                        intAmountInStock3 = Me.txtAmountInStock.Text ' ...this is a numeric value. If not...
                        intCheck3 += 1 ' ... this will not run (check using a breakpoint)...
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid inventory amount") '... and a message box will show
                    End Try

                    Try ' Checks if...
                        decPricePerItem3 = Me.txtPricePerItem.Text ' ...this is a numeric value. If not...
                        intCheck3 += 1 ' ... this will not run (check using a breakpoint)...
                    Catch ex As Exception
                        MessageBox.Show("Enter a valid price") '... and a message box will show
                    End Try


                    decTotalValueInStock3 = decPricePerItem3 * intAmountInStock3 ' Calculates total value of product

                    ' This ensures that the program successfully completed the four statements above...
                    If intCheck3 = 4 Then
                        intButtonPresses += 1 ' and adds 1 to the counter. Max amount of button clicks is three.
                        If intProductNumber1 < intProductNumber2 AndAlso intProductNumber1 < intAmountInStock3 Then ' If Product Number 1 is less than all other products
                            If intProductNumber2 < intProductNumber3 Then ' If the secound product number is lower than the third prodcut number
                                Me.lblProduct1.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct2.Text = intProductNumber2 & vbCrLf & strProductName2 & vbCrLf & intAmountInStock2 & vbCrLf & "$" & decPricePerItem2 & vbCrLf & "$" & decTotalValueInStock2 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct3.Text = intProductNumber3 & vbCrLf & strProductName3 & vbCrLf & intAmountInStock3 & vbCrLf & "$" & decPricePerItem3 & vbCrLf & "$" & decTotalValueInStock3 ' Outputs the data entered by the user from the left side of the program
                            Else ' If product number 2 is not lower then product number 3
                                Me.lblProduct1.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct2.Text = intProductNumber3 & vbCrLf & strProductName3 & vbCrLf & intAmountInStock3 & vbCrLf & "$" & decPricePerItem3 & vbCrLf & "$" & decTotalValueInStock3 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct3.Text = intProductNumber2 & vbCrLf & strProductName2 & vbCrLf & intAmountInStock2 & vbCrLf & "$" & decPricePerItem2 & vbCrLf & "$" & decTotalValueInStock2 ' Outputs the data entered by the user from the left side of the program
                            End If
                        End If

                        If intProductNumber3 < intProductNumber1 AndAlso intProductNumber3 < intProductNumber2 Then ' If Product Number 3 is less then Product Number 1 and Product Number 3 is less then Product Number
                            If intProductNumber2 < intProductNumber1 Then ' If product number 2 is less then product number 1
                                Me.lblProduct1.Text = intProductNumber3 & vbCrLf & strProductName3 & vbCrLf & intAmountInStock3 & vbCrLf & "$" & decPricePerItem3 & vbCrLf & "$" & decTotalValueInStock3 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct2.Text = intProductNumber2 & vbCrLf & strProductName2 & vbCrLf & intAmountInStock2 & vbCrLf & "$" & decPricePerItem2 & vbCrLf & "$" & decTotalValueInStock2 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct3.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                            Else ' If product number 2 is not less then product number 1
                                Me.lblProduct1.Text = intProductNumber3 & vbCrLf & strProductName3 & vbCrLf & intAmountInStock3 & vbCrLf & "$" & decPricePerItem3 & vbCrLf & "$" & decTotalValueInStock3 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct2.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct3.Text = intProductNumber2 & vbCrLf & strProductName2 & vbCrLf & intAmountInStock2 & vbCrLf & "$" & decPricePerItem2 & vbCrLf & "$" & decTotalValueInStock2 ' Outputs the data entered by the user from the left side of the program
                            End If
                        End If
                        If intProductNumber2 < intProductNumber3 AndAlso intProductNumber2 < intProductNumber1 Then ' If product number 2 is less then product number 3 and product number 2 is less then product number 1
                            If intProductNumber3 < intProductNumber1 Then ' If product number 3 is less then product number 1
                                Me.lblProduct1.Text = intProductNumber2 & vbCrLf & strProductName2 & vbCrLf & intAmountInStock2 & vbCrLf & "$" & decPricePerItem2 & vbCrLf & "$" & decTotalValueInStock2 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct2.Text = intProductNumber3 & vbCrLf & strProductName3 & vbCrLf & intAmountInStock3 & vbCrLf & "$" & decPricePerItem3 & vbCrLf & "$" & decTotalValueInStock3 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct3.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                            Else ' If product number 3 is not less then product number 1
                                Me.lblProduct1.Text = intProductNumber2 & vbCrLf & strProductName2 & vbCrLf & intAmountInStock2 & vbCrLf & "$" & decPricePerItem2 & vbCrLf & "$" & decTotalValueInStock2 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct2.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
                                Me.lblProduct3.Text = intProductNumber3 & vbCrLf & strProductName3 & vbCrLf & intAmountInStock3 & vbCrLf & "$" & decPricePerItem3 & vbCrLf & "$" & decTotalValueInStock3 ' Outputs the data entered by the user from the left side of the program
                            End If
                        End If

                    Else
                        intCheck3 = 0 ' If not all operations were completed successfully, then intCheck3 will be set back to 1.
                    End If
                End If
            Case > 2 ' If the user tries to add more then three (var starts at 0) products, a message box will show.
                MessageBox.Show("Maximum of three products met") ' Prevents user from adding more then three products
        End Select

        lblItemNumberSentence.Text = "This concludes the inventory printout. " & intButtonPresses & " item(s) have been displayed. " ' Prints the final display text at the bottom right of the program

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Form
    End Sub

    Private Sub RestartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestartToolStripMenuItem.Click
        Application.Restart() ' Restarts application if selected from strip menu
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit() ' Exits application if selected from strip menu
    End Sub

    Private Sub lblProductNamePromt_Click(sender As Object, e As EventArgs) Handles lblProductNamePromt.Click
        ' Describes the text box to the right.
        ' Says "Product Name"
    End Sub

    Private Sub lblProductNumberPromt_Click(sender As Object, e As EventArgs) Handles lblProductNumberPromt.Click
        ' Describes the text box to the right.
        ' Says "Product Number"
    End Sub

    Private Sub lblPricePerItemPromt_Click(sender As Object, e As EventArgs) Handles lblPricePerItemPromt.Click
        ' Describes the text box to the right.
        ' Says "Price Per Item"
    End Sub

    Private Sub lblBorder_Click(sender As Object, e As EventArgs) Handles lblBorder.Click
        ' Line which divides the Data Entry and Data Review sections.
    End Sub

    Private Sub lblDataEntryTitle_Click(sender As Object, e As EventArgs) Handles lblDataEntryTitle.Click
        ' Says "Data Entry"
    End Sub

    Private Sub lblDataReviewTitle_Click(sender As Object, e As EventArgs) Handles lblDataReviewTitle.Click
        ' Says "Data Review"
    End Sub

    Private Sub lblProductPromt_Click(sender As Object, e As EventArgs) Handles lblProductPromt.Click
        ' Lists the product information categories for each product (one, two, three)
    End Sub

    Private Sub lblItemNumberSentence_Click(sender As Object, e As EventArgs) Handles lblItemNumberSentence.Click
        ' Gives information on how many products have been displayed
        ' Says "This concludes the inventory printout. # items have been displayed.
    End Sub

    Private Sub lblProduct1_Click(sender As Object, e As EventArgs) Handles lblProduct1.Click
        ' List product information for the first product (sorted by item number)
    End Sub

    Private Sub lblProduct2_Click(sender As Object, e As EventArgs) Handles lblProduct2.Click
        ' List product information for the second product (sorted by item number)
    End Sub

    Private Sub lblProduct3_Click(sender As Object, e As EventArgs) Handles lblProduct3.Click
        ' List product information for the third product (sorted by item number)
    End Sub

    Private Sub lblAmountInStockPromt_Click(sender As Object, e As EventArgs) Handles lblAmountInStockPromt.Click
        ' Describes the text box to the right.
        ' Says "Amount In Stock"
    End Sub

    Private Sub txtAmountInStock_TextChanged(sender As Object, e As EventArgs) Handles txtAmountInStock.TextChanged
        ' Textbox which the user uses to input Amount In Stock
    End Sub

    Private Sub txtProductNumber_TextChanged(sender As Object, e As EventArgs) Handles txtProductNumber.TextChanged
        ' Textbox which the user uses to input Product Number
    End Sub

    Private Sub txtProductName_TextChanged(sender As Object, e As EventArgs) Handles txtProductName.TextChanged
        ' Textbox which the user uses to input Product Name
    End Sub

    Private Sub txtPricePerItem_TextChanged(sender As Object, e As EventArgs) Handles txtPricePerItem.TextChanged
        ' Textbox which the user uses to input Price Per Item
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked
        ' Strip Menu
    End Sub
End Class


' While making this program, I reinforced my knowledge of using nested If..Then..Else statements 
' and using breakpoints while debugging.

' While making this program, I had difficulties with a situation were my program would reset some 
' of the variables leading to incorrect output data. I fixed this by making the variables global.