' Honors Inventory Project Part III
' Jonathan Rogers
' 8/13/2020

' The purpose of this program is to organize product information entered by the user into a Listbox. 
' This program collects information such as product name, product number, the amount in stock, and the 
' price per item. Using this information, the program both organizes the products, but also calculates 
' the total value of each product (in-stock only), the number of products being displayed, and 
' the total value of all items in stock among other pieces of information,

' This program uses four labels, one button, a list box, and one strip menu on a Windows Form.

Public Class Form1

    ' These variables are arrays which store the product information for each group based on their name
    Dim strProductName() As String ' Stores product name for all products
    Dim intProductNumber() As Integer ' Stores product number for all products
    Dim intAmountInStock() As Integer ' Stores amount in stock for all products
    Dim decPricePerItem() As Decimal ' Stores price per item for all products
    Dim decInventoryValue() As Decimal ' Stores the total value of the items in stock for all products.


    Dim intNumberOfItems As Integer ' Stores the number of items that will be entered by the user

    Dim intMaxInventoryValue As Integer ' Stores the max inventory value out of all the items entered by the user
    Dim strMaxInventoryValueProductName As String ' Stores the item name which parallels to the max inventory value



    ' *** Sub procedures and functions ***


    Public Sub DataToList() ' Adds products to list
        ' Pre: Optains the product number, product name, amount in stock, price per item, and value in stock
        ' and puts them into the list box.
        Dim intCounter As Integer

        Do While intCounter < intNumberOfItems

            If strProductName(intCounter) = Nothing Then ' If there is no product name for a current item, the program will exit this procedure. This prevents the problem of a long list of groups of products which had no data because the user canceled data entery early.
                Exit Sub
            End If

            lstData.Items.Add("---------------------------------------------------------------------------------------------------------------------") ' Barrier (for aesthetics)
            lstData.Items.Add("Item Number:" & vbTab & vbTab & intProductNumber(intCounter)) ' Item number
            lstData.Items.Add("Product Name:" & vbTab & vbTab & strProductName(intCounter)) ' Product Name
            lstData.Items.Add("Items in stock:" & vbTab & vbTab & intAmountInStock(intCounter)) ' Items in stock
            lstData.Items.Add("Price of item:" & vbTab & vbTab & "$" & decPricePerItem(intCounter)) ' Price of item
            lstData.Items.Add("Total value of item in stock:" & vbTab & "$" & decInventoryValue(intCounter)) ' Total value of items in stock

            intCounter += 1
        Loop


        ' Post: Product information will be displayed on the listbox.
    End Sub


    Public Function inventoryValue(Price, Amount)
        ' Pre: This function optains the price and in stock amount of the current product being calculated.
        ' It will then complete calculations to determine the total value in stock for the item.

        ' ** Product is used to describe the mathematical operation not the item **

        Dim decTotal As Decimal ' Stores the product of the price of an item and the quantity of the item.

        decTotal = Price * Amount ' Calculates the value of the product (in stock only)
        Return decTotal ' Returns the product where the function was called.


        ' Post: Returns the total inventory value.
    End Function

    Public Function TotalValue()
        ' Pre: This function optains the inventory value of each item entered which is added together to get a total value.

        Dim intCounter As Integer ' Counts the amount of times the loop runs
        Dim decTotal As Decimal ' Stores the total value of all inventory

        Do While intCounter < intNumberOfItems
            decTotal += decInventoryValue(intCounter) ' Adds the inventory value of the current product (product determined by intCounter) to the total value
            intCounter += 1 ' Ensures loop does not run forever.
        Loop

        Return decTotal ' Returns the total value for use outside the function

        ' Post: Returns the total inventory value of all items
    End Function

    Public Sub MostValuableItem()
        ' Pre: This procedure optains the max value from decInventoryValue which is used to define a variable to be used elsewhere in the program.
        ' This program also finds the parallel element (decInventoryValue --- strProductName) by geting the index of the max value (of decInventoryValue) 
        ' to then get the correct the element from strProductName.

        intMaxInventoryValue = decInventoryValue.Max ' Finds max inventory value from the array (using the .Max class) and sets it equal to intMaxInventoryValue
        strMaxInventoryValueProductName = strProductName(Array.IndexOf(decInventoryValue, decInventoryValue.Max)) ' Finds the index of the largest value and uses such index to get the parallel element from strProductName

        ' Post: Sets both values to a global variable to be used elsewhere in the program.
    End Sub



    ' *** Event Procedures ***


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

        ' Local Variables

        Dim intCounter As Integer ' Used to count the amount of times the Do While Loop runs
        Dim intCheck As Integer ' Checks to ensure if all values were successfully entered
        Dim strNumberOfItems As String ' Var which will store the amount of products the user will put into the program



        Try
            If lstData.Items.Contains("---------------------------------------------------------------------------------------------------------------------") Then ' If a set of values has already been entered in...
                If MessageBox.Show("Updating data will remove all preexisting data. Are you sure you want to continue?", "Update Data", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then ' Ask the user if they would like to overwrite data to make a new set
                    Exit Try ' If they don't, exit try
                End If

            End If
            strNumberOfItems = InputBox("Insert the amount of products you wish to display.", "Product Amount") ' Asks user to input amount of products to be entered
            If strNumberOfItems = Nothing Then
                Exit Try
            End If
            intNumberOfItems = Convert.ToInt32(strNumberOfItems) ' Converts the string value to an integer

            ' ReDim's arrays to match with the amount of products the user will enter
            ReDim strProductName(intNumberOfItems)
            ReDim intProductNumber(intNumberOfItems)
            ReDim intAmountInStock(intNumberOfItems)
            ReDim decPricePerItem(intNumberOfItems)
            ReDim decInventoryValue(intNumberOfItems)


            lstData.Items.Clear() ' Clears items in listbox for next loop
            lblItemNumberSentence.Text = "This concludes the inventory printout. ____ items have been displayed for a total inventory" & vbCrLf & "value of $_____. __ .  The most valuable item in the inventory is _______ at a value of $_____. __ ." ' Resets ending phrase
            Do While intCounter < intNumberOfItems ' Loop while intCounter is less then the number of items that the user will enter

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
                            Exit Sub ' ...the program will exit the Loop.
                        End If
                        intProductNumber(intCounter) = Convert.ToInt32(strInputValue) ' Converts string value to integer

                        intCheck += 1 ' One is added to intCheck
                    Catch Ex As Exception
                        MessageBox.Show("Enter a numeric value.", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                        Exit Sub
                    End Try

                    ' Product in stock
                    Try
                        strInputValue = InputBox("Input amount in stock for product " & intCounter + 1) ' Shows an input box and assigns the input to a variable.
                        If strInputValue = Nothing Then ' If the value from the input box above is equal to nothing...
                            intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                            Exit Sub ' ...the program will exit the Loop.
                        End If
                        intAmountInStock(intCounter) = Convert.ToInt32(strInputValue) ' Converts string value to integer
                        intCheck += 1
                    Catch ex As Exception
                        MessageBox.Show("Enter a numeric value.", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                        Exit Sub
                    End Try

                    ' Price per item
                    Try
                        strInputValue = InputBox("Input price per item for product " & intCounter + 1) ' Shows an input box and assigns the input to a variable.
                        If strInputValue = Nothing Then ' If the value from the input box above is equal to nothing...
                            intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                            Exit Sub ' ...the program will exit the Loop.
                        End If
                        decPricePerItem(intCounter) = Convert.ToDecimal(strInputValue) ' Converts string value to integer
                        intCheck += 1
                    Catch ex As Exception
                        MessageBox.Show("Enter a numeric value.")
                        intCheck = 0 ' Because this section failed, intCheck will go back to zero as the process starts over again
                        Exit Sub
                    End Try

                    decInventoryValue(intCounter) = inventoryValue(decPricePerItem(intCounter), intAmountInStock(intCounter)) ' Runs a function which calculates the total value of the products.


                    intCounter += 1 ' Adds 1 to intCounter to ensure program does not go over a user-specifed amount. This also keeps a record of how many prodcuts are been entered.
                    intCheck = 0 ' Sets intCheck to zero for use in the next loop

                End If

            Loop

            Call MostValuableItem() ' Finds the most valuable item and the corresponding item name.
            Call DataToList() ' Calls DataToList which adds all the data to the list box
            lblItemNumberSentence.Text = "This concludes the inventory printout. " & intCounter & " items have been displayed for a total inventory" & vbCrLf & "value of $" & TotalValue() & ". The most valuable item in the inventory is " & strMaxInventoryValueProductName & " at a value of $" & intMaxInventoryValue & "." ' Adds required information to the ending text

        Catch ex As Exception
            MessageBox.Show("An error has occurred. Check your values and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) ' If an error occurs, a message box will display
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


' While making this program, I reinforced my knowledge of using the math class, arrays, procedures, and functions.

' While making this program, I had two main difficulties. This one difficulty was using the .Max class to get the largest 
' value from an array. I tried several ideas before I found the solution that I believe works best in this situation.
' The other difficulties occurred when the user would tell the program that they would want to enter x-amount of items but 
' only entered a portion of that, the program would display both the items entered and items that were not entered into the 
' list box. This was fixed by using an If statement before the data from the arrays were added to the Listbox.