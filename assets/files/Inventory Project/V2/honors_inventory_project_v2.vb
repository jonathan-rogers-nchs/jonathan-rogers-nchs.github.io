' Honors Inventory Project Part 2
' Jonathan Rogers
' 8/6/2020

' The purpose of this program is to organize product information entered by the user into a Listbox. 
' This program collects information such as product name, product number, the amount in stock, and the 
' price per item. Using this information, the program both organizes the products, but also calculates 
' the total value of each product (in-stock only), the number of products being displayed (max 5), and 
' the total value of all items in stock.

' This program uses four labels, one button, a list box, and one strip menu on a Windows Form.

Public Class Form1

    ' These variables are arrays which store the product information for each group based on their name
    Dim strProductName(4) As String ' Stores product name for all products
    Dim intProductNumber(4) As Integer ' Stores product number for all products
    Dim intAmountInStock(4) As Integer ' Stores amount in stock for all products
    Dim decPricePerItem(4) As Decimal ' Stores price per item for all products
    Dim decValueInStock(4) As Decimal ' Stores the total value of the items in stock for all products.



    ' The following lines of procedures store the algorithm to print (output) for each product when called.
    Public Sub Product1() ' First product
        ' Optains the product number, product name, amount in stock, price per item, and value in stock
        ' and puts them into the list box.

        lstData.Items.Add("---------------------------------------------------------------------------------------------------------------------") ' Barrier
        lstData.Items.Add("Item Number:" & vbTab & vbTab & intProductNumber(0)) ' Item number
        lstData.Items.Add("Product Name:" & vbTab & vbTab & strProductName(0)) ' Product Name
        lstData.Items.Add("Items in stock:" & vbTab & vbTab & intAmountInStock(0)) ' Items in stock
        lstData.Items.Add("Price of item:" & vbTab & vbTab & "$" & decPricePerItem(0)) ' Price of item
        lstData.Items.Add("Total value of item in stock:" & vbTab & "$" & decValueInStock(0)) ' Total value of items in stock

        ' Post: Product information will be displayed on the listbox.
    End Sub

    Public Sub Product2() ' Second product
        ' Optains the product number, product name, amount in stock, price per item, and value in stock
        ' and puts them into the list box.

        lstData.Items.Add("---------------------------------------------------------------------------------------------------------------------") ' Barrier
        lstData.Items.Add("Item Number:" & vbTab & vbTab & intProductNumber(1)) ' Item number
        lstData.Items.Add("Product Name:" & vbTab & vbTab & strProductName(1)) ' Product Name
        lstData.Items.Add("Items in stock:" & vbTab & vbTab & intAmountInStock(1)) ' Items in stock
        lstData.Items.Add("Price of item:" & vbTab & vbTab & "$" & decPricePerItem(1)) ' Price of item
        lstData.Items.Add("Total value of item in stock:" & vbTab & "$" & decValueInStock(1)) ' Total value of items in stock

        ' Post: Product information will be displayed on the listbox.
    End Sub

    Public Sub Product3() ' Third product
        ' Optains the product number, product name, amount in stock, price per item, and value in stock
        ' and puts them into the list box.

        lstData.Items.Add("---------------------------------------------------------------------------------------------------------------------") ' Barrier
        lstData.Items.Add("Item Number:" & vbTab & vbTab & intProductNumber(2)) ' Item number
        lstData.Items.Add("Product Name:" & vbTab & vbTab & strProductName(2)) ' Product Name
        lstData.Items.Add("Items in stock:" & vbTab & vbTab & intAmountInStock(2)) ' Items in stock
        lstData.Items.Add("Price of item:" & vbTab & vbTab & "$" & decPricePerItem(2))
        lstData.Items.Add("Total value of item in stock:" & vbTab & "$" & decValueInStock(2)) ' Total value of items in stock

        ' Post: Product information will be displayed on the listbox.
    End Sub

    Public Sub Product4() ' Fourth product
        ' Optains the product number, product name, amount in stock, price per item, and value in stock
        ' and puts them into the list box.

        lstData.Items.Add("---------------------------------------------------------------------------------------------------------------------") ' Barrier
        lstData.Items.Add("Item Number:" & vbTab & vbTab & intProductNumber(3)) ' Item number
        lstData.Items.Add("Product Name:" & vbTab & vbTab & strProductName(3)) ' Product Name
        lstData.Items.Add("Items in stock:" & vbTab & vbTab & intAmountInStock(3)) ' Items in stock
        lstData.Items.Add("Price of item:" & vbTab & vbTab & "$" & decPricePerItem(3)) ' Price of item
        lstData.Items.Add("Total value of item in stock:" & vbTab & "$" & decValueInStock(3)) ' Total value of items in stock

        ' Post: Product information will be displayed on the listbox.
    End Sub

    Public Sub Product5() ' Fifth product
        ' Optains the product number, product name, amount in stock, price per item, and value in stock
        ' and puts them into the list box.

        lstData.Items.Add("---------------------------------------------------------------------------------------------------------------------") ' Barrier
        lstData.Items.Add("Item Number:" & vbTab & vbTab & intProductNumber(4)) ' Item number
        lstData.Items.Add("Product Name:" & vbTab & vbTab & strProductName(4)) ' Product Name
        lstData.Items.Add("Items in stock:" & vbTab & vbTab & intAmountInStock(4)) ' Items in stock
        lstData.Items.Add("Price of item:" & vbTab & vbTab & "$" & decPricePerItem(4)) ' Price of item
        lstData.Items.Add("Total value of item in stock:" & vbTab & "$" & decValueInStock(4)) ' Total value of items in stock

        ' Post: Product information will be displayed on the listbox.
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

        ' Local Variables

        Dim intNumberOfItems As Integer ' Stores user input after converted to Integer
        Dim intCounter As Integer ' Used to count the amount of times the Do While Loop runs
        Dim intCheck As Integer ' Checks to ensure if all values were successfully entered
        Dim strNumberOfItems As String ' Var which will store the amount of products the user will put into the program



        Try
            If lstData.Items.Contains("---------------------------------------------------------------------------------------------------------------------") Then ' If a set of values has already been entered in...
                If MessageBox.Show("Updating data will remove all preexisting data. Are you sure you want to continue?", "Update Data", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then ' Ask the user if they would like to overwrite data to make a new set
                    Exit Try ' If they don't, exit try
                End If

            End If
            strNumberOfItems = InputBox("Insert the amount of products you wish to display from 1 to 5.", "Product Amount") ' Asks user to input amount of products to be entered
            If strNumberOfItems = Nothing Then
                Exit Try
            End If
            intNumberOfItems = Convert.ToInt32(strNumberOfItems) ' Converts the string value to an integer
            If intNumberOfItems > 5 Or intNumberOfItems < 1 Then ' Test if value is between 1 and 5
                MessageBox.Show("Value must be from 1 to 5") ' Warns user that the input must be between 1 and 5
                Exit Try ' Exits Try
            End If
            lstData.Items.Clear() ' Clears items in listbox for next loop
            lblItemNumberSentence.Text = "This concludes the inventory printout. ____ items have been displayed for a total inventory value " & vbCrLf & " of $_____ . __ " ' Resets ending phrase
            Do While intCounter < intNumberOfItems ' Loop while intCounter is less then the number of items that the user will enter


                ' Section 1: Getting data from user and calculations

                strProductName(intCounter) = InputBox("Input product name for product " & intCounter + 1) ' Shows an input box asking for the name of the product

                If strProductName(intCounter) = Nothing Then ' If no name is entered...
                    Exit Do ' ...the loop will exit
                Else
                    Dim strInputValue As String ' This stores the value from the inputbox below before it is sent to the proper variable. This allows the program to act if cancel is selected.

                    ' Product Number
                    Try
                        strInputValue = InputBox("Input product number for product " & intCounter + 1) ' Shows an input box and assigns the input to a variable.
                        If strInputValue = Nothing Then ' If the value from the input box above is equal to nothing...
                            intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                            Exit Do ' ...the program will exit the Loop.
                        End If
                        intProductNumber(intCounter) = Convert.ToInt32(strInputValue) ' Converts string value to integer

                        intCheck += 1 ' One is added to intCheck
                    Catch Ex As Exception
                        MessageBox.Show("Enter a numeric value.")
                        intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                        Exit Do
                    End Try

                    ' Product in stock
                    Try
                        strInputValue = InputBox("Input amount in stock for product " & intCounter + 1) ' Shows an input box and assigns the input to a variable.
                        If strInputValue = Nothing Then ' If the value from the input box above is equal to nothing...
                            intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                            Exit Do ' ...the program will exit the Loop.
                        End If
                        intAmountInStock(intCounter) = Convert.ToInt32(strInputValue) ' Converts string value to integer
                        intCheck += 1
                    Catch ex As Exception
                        MessageBox.Show("Enter a numeric value.")
                        intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                        Exit Do
                    End Try

                    ' Price per item
                    Try
                        strInputValue = InputBox("Input price per item for product " & intCounter + 1) ' Shows an input box and assigns the input to a variable.
                        If strInputValue = Nothing Then ' If the value from the input box above is equal to nothing...
                            intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                            Exit Do ' ...the program will exit the Loop.
                        End If
                        decPricePerItem(intCounter) = Convert.ToDecimal(strInputValue) ' Converts string value to integer
                        intCheck += 1
                    Catch ex As Exception
                        MessageBox.Show("Enter a numeric value.")
                        intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                        Exit Do
                    End Try

                    decValueInStock(intCounter) = decPricePerItem(intCounter) * intAmountInStock(intCounter) ' Calculates the total value of the products in stock



                    ' Section 2: Finding the lowest number and displaying it



                    Dim intLowestProductNumber As Integer ' Stores the lowest product number to determine which product should fo first

                    ' The next series of If..Else statements are used to determine the product with the lowest product number.
                    ' The product with the lowest value will then be placed first compared to the other products.

                    If intCounter = 0 Then '  If intCounter is equal to zero (if the loop is running for the first time)

                        Call Product1() ' The first product will be added to the list box
                        intLowestProductNumber = intProductNumber(0) ' The lowest product number will be set to the product number of product 1
                    End If

                    If intCounter = 1 Then
                        If intProductNumber(intCounter) > intLowestProductNumber Then '  If intCounter is equal to one (if the loop is running for the second time) and the current product number is greator then the current lowest number
                            Call Product2() ' The second product will be added to the list box
                        Else
                            Me.lstData.Items.Clear() ' All items will be cleared
                            ' The following products will be added
                            Call Product2()
                            Call Product1()
                            intLowestProductNumber = intProductNumber(intCounter) ' The lowest product number will be set to the product number of product 2
                        End If
                    End If

                    If intCounter = 2 Then
                        If intProductNumber(intCounter) > intLowestProductNumber Then '  If intCounter is equal to two (if the loop is running for the third time) and the current product number is greator then the current lowest number
                            Call Product3() ' The third product will be added to the list box
                        Else
                            Me.lstData.Items.Clear() ' All items will be cleared
                            ' The following products will be added
                            Call Product3()
                            Call Product1()
                            Call Product2()
                            intLowestProductNumber = intProductNumber(intCounter) ' The lowest product number will be set to the product number of product 3
                        End If
                    End If

                    If intCounter = 3 Then
                        If intProductNumber(intCounter) > intLowestProductNumber Then '  If intCounter is equal to three (if the loop is running for the fourth time) and the current product number is greator then the current lowest number
                            Call Product4() ' The fourth product will be added to the list box
                        Else
                            Me.lstData.Items.Clear() ' All items will be cleared
                            ' The following products will be added
                            Call Product4()
                            Call Product1()
                            Call Product2()
                            Call Product3()
                            intLowestProductNumber = intProductNumber(intCounter) ' The lowest product number will be set to the product number of product 4
                        End If
                    End If

                    If intCounter = 4 Then
                        If intProductNumber(intCounter) > intLowestProductNumber Then '  If intCounter is equal to four (if the loop is running for the fifth time) and the current product number is greator then the current lowest number
                            Call Product5() ' The fifth product will be added to the list box
                        Else
                            Me.lstData.Items.Clear() ' All items will be cleared
                            ' The following products will be added
                            Call Product5()
                            Call Product1()
                            Call Product2()
                            Call Product3()
                            Call Product4()
                            intLowestProductNumber = intProductNumber(intCounter) ' The lowest product number will be set to the product number of product 5
                        End If
                    End If

                    intCounter += 1 ' Adds 1 to intCounter to ensure program does not go over a pre-specifed amount (5 runs)
                    intCheck = 0 ' Sets intCheck to zero for use in the next loop

                    ' Section 3: Calculating total product value and products

                    Dim decTotalValue As Decimal ' Stores total value of all items
                    decTotalValue = decValueInStock(0) + decValueInStock(1) + decValueInStock(2) + decValueInStock(3) + decValueInStock(4) ' Finds the sum of product value for each product

                    lblItemNumberSentence.Text = "This concludes the inventory printout. " & intCounter & " item(s) have been displayed for a total inventory value " & vbCrLf & " of $" & decTotalValue & "." ' End string summarizing presented data
                End If

            Loop
        Catch ex As Exception
            MessageBox.Show("Numeric values only") ' If an error occurs, a message box will display saying only numeric values are allowed
        End Try



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

    Private Sub lblBorder_Click(sender As Object, e As EventArgs) Handles lblBorder.Click
        ' Line which divides the Data Entry and Data Review sections.
    End Sub

    Private Sub lblDataEntryTitle_Click(sender As Object, e As EventArgs) Handles lblDataEntryTitle.Click
        ' Says "Data Entry"
    End Sub

    Private Sub lblDataReviewTitle_Click(sender As Object, e As EventArgs) Handles lblDataReviewTitle.Click
        ' Says "Data Review"
    End Sub


    Private Sub lblItemNumberSentence_Click(sender As Object, e As EventArgs) Handles lblItemNumberSentence.Click
        ' Gives information on how many products have been displayed
        ' Says "This concludes the inventory printout. # items have been displayed.
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked
        ' Strip Menu
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstData.SelectedIndexChanged
        ' List Box
    End Sub

    Private Sub ClearAllItemsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllItemsToolStripMenuItem.Click
        lstData.Items.Clear() ' Clears all content in list box
    End Sub
End Class


' While making this program, I reinforced my knowledge of using arrays, procedures, and list boxes.

' While making this program, I had some difficulties with different parts of this project. First, 
' I had difficulties creating a system were the item with the lowest product number would come 
' first while the rest followed. I solved this by using procedures to quickly add data to the list 
' box. The next part of the project I had struggled with was assigning variables. I wanted to use a 
' loop structure to automatically assign values to each variable but before learning about arrays, I 
' did not know how to do this. Now, with my knowledge of arrays, I was able to successfully use a loop 
' to assign values to variables.