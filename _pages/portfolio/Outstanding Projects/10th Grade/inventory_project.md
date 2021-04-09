---
permalink: /portfolio/outstanding-projects/inventory-project/
title: "Inventory Project"
author_profile: false
layout: splash
header:
  overlay_image: /assets/images/outstanding_projects-header.jpg 
  overlay_filter: 0.5
  caption: "Photo credit: [**'Startup Stock Photos' on Pexels**](https://www.pexels.com/photo/blue-printer-paper-7376/)"
toc: true
toc_label: " Table of Contents"
toc_icon: "file-alt"
---

***Special Note:** The contents of the website page includes vocabulary and/or advanced computer concepts that may be unclear to the majority of people. To fully understand the contents of the page, please read everything including the code within the code blocks. More information may be also found on the official Microsoft Documentation.*

<a href="https://docs.microsoft.com/en-us/dotnet/visual-basic/" class="btn btn--inverse btn--x-large">Official Microsoft Documentation</a>

# Overview
During the summer of the 2019-2020 school year, I enrolled in Computer Programing I Honors as a way to improve my programming abilities in software development. As this was an honors course, it was required that an honors project was completed and submitted to demonstrate student growth during the class. The topic for this project was inventory management.

# Inventory Project Part I
This project had three major portions, each of which would be added on to the last to create one final program. The first portion of this project was to show a student's ability in the use of Controls, Properties, Variables, Scope, Text boxes, Message Boxes, Calculations, Errors, Debugging, Conditional Statements, Math, Boolean and Relational operators, Random, Static, and Counters.

I started this portion of the project by adding all of the global variables that I would be using. These included variables that I would use to store the product name, number, price, etc.

~~~~
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
~~~~

I then added a button called "btnUpdate" that when pressed the data would be received, validated, and displayed. In the button code, I added local variables which ensures the integrity of data

~~~~
Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

        ' Local Variables
        Static intButtonPresses As Integer ' Stores the amount of successful button clicks

        Static intCheck1 As Integer = 0 ' Stores the amount of successful operations completed by computer for product 1. 
        Static intCheck2 As Integer = 0 ' Stores the amount of successful operations completed by computer for product 2.
        Static intCheck3 As Integer = 0 ' Stores the amount of successful operations completed by computer for product 3.
~~~~

This program runs by first checking how many times the button has been clicked (and data shown on the program). This is done by using a Select Case statement to check if intButtonPresses is equal to a certain value. For example, if intButtonPresses is equal to 1, the code under Case = 1 will run.

~~~~
Select Case intButtonPresses
            Case = 0
                ' Some code
                intButtonPresses += 1 ' and adds 1 to the counter.
            Case = 1
                ' Some more code
                intButtonPresses += 1 ' and adds 1 to the counter.
            Case 2
                ' Even more code
                intButtonPresses += 1 ' and adds 1 to the counter.
            Case > 2 ' If the user tries to add more then three (var starts at 0) products, a message box will show.
                MessageBox.Show("Maximum of three products met") ' Prevents user from adding more then three products
        End Select
~~~~

After the program has determined what code will run in the Select Case statement, the program will go through a series of steps to ensure the data is usable. First, the program will get a value that the user entered into the textbox. If there is no data in the textbox, the user will get a message saying "Enter a product name." If this value is not true, then the program will continue to check the rest of the entered values.

~~~~
strProductName1 = Me.txtProductName.Text ' Sets the values in txtProductName to a variable
                If strProductName1 = Nothing Then ' If there is no values...
                    MessageBox.Show("Enter a product name") ' ... a message box will display saying "Enter a product name"
                Else ' If there is text...
~~~~

Now the program will check if the entered product number, the amount in stock value, and price per item value are integers or rational numbers. The program does this by using a Try statement. The Try statement checks if any code that it is checking returns a runtime error (an error that occurs while the program is running). If the Try statement does get an error, the program will not crash but will display a warning to the user describing the problem. The program will also not run the rest of the code below it (within Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click).

~~~~
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
~~~~

Finally, the program will make the needed calculations and display the results on the UI of the program.

~~~~
decTotalValueInStock1 = decPricePerItem1 * intAmountInStock1 ' Calculates total value of product

' This ensures that the program successfully completed the four statements above...
If intCheck1 = 4 Then ' If all operations are completed correctly
    intButtonPresses += 1 ' and adds 1 to the counter. Max amount of button clicks is three.
    Me.lblProduct1.Text = intProductNumber1 & vbCrLf & strProductName1 & vbCrLf & intAmountInStock1 & vbCrLf & "$" & decPricePerItem1 & vbCrLf & "$" & decTotalValueInStock1 ' Outputs the data entered by the user from the left side of the program
Else
    intCheck1 = 0 ' If not all operations were completed successfully, then intCheck1 will be set back to 1.
End If
~~~~

In other parts of the program, for example, if the user has already entered some data, the program will automatically sort (by product number) and add the new data as follows.

~~~~
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
~~~~

The final program looks like this...

<figure class="single">
    <a href="/assets/images/Inventory Project/Project V1.png"><img src="/assets/images/Inventory Project/Project V1.png"></a>
</figure>

The project files can be found below.

<a href="https://github.com/jonathan-rogers-dev/Inventory-Manager/releases/download/V1/honors_inventory_project_v1.exe" class="btn btn--inverse btn--x-large">Executable (.exe)</a>
<a href="/assets/files/Inventory Project/V1/honors_inventory_project.vb" class="btn btn--inverse btn--x-large">Form (.vb)</a>


## Inventory Project Part I Reflection
As a new programmer, I am very happy with the results of this first program. This program is easy to use, simple, and robust. By the use of data validation, this program will not crash in most circumstances if the user inputs invalid data but will warn them of the error. I am also very satisfied with the sorting algorithm which ensures the program is placed in order by product number. Although I did have many successes, I learned a lot from making this program. I learned that in order to prevent a variable from being redeclared (reinitialized) every time the button was clicked (redeclaring a variable resets its value to 0 or ""), I need to make the variable global meaning it is initialized when the program starts. I also used the built-in debugging tool to fix some problems I ran into during development. While I am greatly satisfied with the results of this program, I found that I could not solve this one problem. This program only allows a max of three products which in a lot of real-world cases, is not sufficient. In order to add more products, I would need to create a lot more variables for each product and be able to print (display) data on the program. With the knowledge I have right now, I can't do that efficiently.

# Inventory Project Part II
For the next part of this project, we were instructed to reprogram portions of the program to fit the additional criteria.
 - Use Input Boxes (with validation) to get user input
 - Display results in a list box
 - Calculate the total value of the inventory
In order to do this, several changes will need to be made to the program.

The first change that needed to be changed was how the user imputed data. Originally, the user would enter the data into textboxes, select "Update" and the data would be displayed on a label. With the new requirements, I would need to use an input box, a pop-up box allowing the user to input information and get user data. Here is an example of how an input box was used. In this example, the user would be shown an input box asking the user to "Insert the number of products you wish to display from 1 to 5."

~~~~
strNumberOfItems = InputBox("Insert the amount of products you wish to display from 1 to 5.", "Product Amount") ' Asks user to input amount of products to be entered
~~~~

Next, the data that would be visible to the user needed to be placed in a list box rather than a label. The following code is an example of how I used Items. Add to add information to a list box after the data has been entered by the user.

~~~~
lstData.Items.Add("---------------------------------------------------------------------------------------------------------------------") ' Barrier
lstData.Items.Add("Item Number:" & vbTab & vbTab & intProductNumber(0)) ' Item number
lstData.Items.Add("Product Name:" & vbTab & vbTab & strProductName(0)) ' Product Name
lstData.Items.Add("Items in stock:" & vbTab & vbTab & intAmountInStock(0)) ' Items in stock
lstData.Items.Add("Price of item:" & vbTab & vbTab & "$" & decPricePerItem(0)) ' Price of item
lstData.Items.Add("Total value of item in stock:" & vbTab & "$" & decValueInStock(0)) ' Total value of items in stock
~~~~

The final change that needed to be made was displaying the total value of the inventory. To do this, I compiled all of the product values together and set the sum to a new variable named decTotalValue. After the sum was found, the program would display that data in a sentence using a label.

~~~~
' Section 3: Calculating total product value and products

Dim decTotalValue As Decimal ' Stores total value of all items
decTotalValue = decValueInStock(0) + decValueInStock(1) + decValueInStock(2) + decValueInStock(3) + decValueInStock(4) ' Finds the sum of product value for each product

lblItemNumberSentence.Text = "This concludes the inventory printout. " & intCounter & " item(s) have been displayed for a total inventory value " & vbCrLf & " of $" & decTotalValue & "." ' End string summarizing presented data
~~~~

In addition to the required changes, I also made some additional changes that would both, make the program easier to read, update, and to make the program more efficient. The first change that was made was to replace a majority of variables with arrays. Arrays are similar to variables, but an array can store multiple different values whereas variables can only store one. These values can then be accessed by using an index value. This would allow the program to both accept more values (this is shown better in the next section), and shorten the length of the program.

~~~~
' These variables are arrays which store the product information for each group based on their name
Dim strProductName(4) As String ' Stores product name for all products
Dim intProductNumber(4) As Integer ' Stores product number for all products
Dim intAmountInStock(4) As Integer ' Stores amount in stock for all products
Dim decPricePerItem(4) As Decimal ' Stores price per item for all products
Dim decValueInStock(4) As Decimal ' Stores the total value of the items in stock for all products.
~~~~

Finally, I changed how the program displayed the data on the list box. Originally, the program used a very long system for sorting and displaying data as seen in the first portion of this project ("Sorting Data" code block). To make this program shorter, I used what is called a procedure. A procedure stores a set of code that will only be executed when it is called. I used multiple procedures to display the data in the list box.

~~~~
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
~~~~

~~~~
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
~~~~

Other changes made to the program
 - All Case If statements replaced with an If..Else statement
 - Stronger data validation

The final program looks like this...

<figure class="single">
    <a href="/assets/images/Inventory Project/Project V2.png"><img src="/assets/images/Inventory Project/Project V2.png"></a>
</figure>

The project files can be found below.

<a href="https://github.com/jonathan-rogers-dev/Inventory-Manager/releases/download/V2/honors_inventory_project_v2.exe" class="btn btn--inverse btn--x-large">Executable (.exe)</a>
<a href="/assets/files/Inventory Project/V2/honors_inventory_project_v2.vb" class="btn btn--inverse btn--x-large">Form (.vb)</a>

## Inventory Project Part II Reflection
After completing the second portion of the inventory project, I was able to apply my knowledge of using arrays, procedures, and list boxes and incorporate those skills into a useful program. Like before, this program is simple and clean, and now, even more, redundant to errors than before. I also learned a lot while making this project. I learned how to use a loop to assign values to an array with the help of a counter and how to use a procedure. For myself, sorting data was one of the most difficult portions of this project. To fix this, I used an If..Else statements similar to ones used in the first part of the project to sort the data. I now realize after programing all three portions of the project that there was an even easier and more efficient method to sort the data. I am also not satisfied with the limit on the product size but I believe with a little more knowledge, I will be able to solve this problem.

# Inventory Project Part III
For the final portion of the inventory project, we were required to complete the following tasks.

 - Using an input box with appropriate prompts and choices of OK or Cancel ask the user for the number of items they wish to enter. Prompt the user what to do if they do not wish to enter any items.create your arrays to be this size.
 - Handle if they click OK without entering anything, they click Cancel without entering anything or they enter an inappropriate number/value.
 - Use at least one loop of your choice, subprocedure and function, and other necessary structures to:
 - Create and call a procedure/function that takes in user input, validates and places the data in parallel arrays
 - Create and call a procedure/function with parameters (quantity on hand and price) that calculates the value of each inventory item and places that value in an additional parallel array called inventoryValue
 - Create and call a procedure/function that finds the most valuable item in the inventory - use a method from the Math Class. (Links to an external site.)
 - Create and call a procedure/function that calculates the total value of the inventory (hint: use inventoryValue to do this)
 - Keep a count of the total number of inventory products – just the different items, not each item in stock.
 - Create and call a procedure/function that displays all inventory items including categories in a pleasing manner similar to the output below in one List Box. Remember to clear output between runs.

These changes would drastically change the internal code of the program but will add many features for the user such as.
 - Large potential datasets
 - Improved UI
 - Faster processing

The first change that needed to be made was the size of the arrays. Currently, they are set at four but we want the array size to change based on user input. In order to do this, the arrays will need to be redeclared after the user defines how many items the program will process.

~~~~
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
~~~~

Now the user can define a specific amount of values that the program will process. Next, a loop will need to be created that will allow the user to enter the product information for each product.

~~~~
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
~~~~

Directly following the loop, three procedures (two sub procedures and one function) will be called to calculate the most valuable item, find the total value of the products, and send all the data to the visible list box.

~~~~
Call MostValuableItem() ' Finds the most valuable item and the corresponding item name.
Call DataToList() ' Calls DataToList which adds all the data to the list box
lblItemNumberSentence.Text = "This concludes the inventory printout. " & intCounter & " items have been displayed for a total inventory" & vbCrLf & "value of $" & TotalValue() & ". The most valuable item in the inventory is " & strMaxInventoryValueProductName & " at a value of $" & intMaxInventoryValue & "." ' Adds required information to the ending text
~~~~

Each Call keyword informs the program to run a separate sub procedure/function.

~~~~
Public Sub MostValuableItem()
    ' Pre: This procedure optains the max value from decInventoryValue which is used to define a variable to be used elsewhere in the program.
    ' This program also finds the parallel element (decInventoryValue --- strProductName) by geting the index of the max value (of decInventoryValue) 
    ' to then get the correct the element from strProductName.

    intMaxInventoryValue = decInventoryValue.Max ' Finds max inventory value from the array (using the .Max class) and sets it equal to intMaxInventoryValue
    strMaxInventoryValueProductName = strProductName(Array.IndexOf(decInventoryValue, decInventoryValue.Max)) ' Finds the index of the largest value and uses such index to get the parallel element from strProductName

    ' Post: Sets both values to a global variable to be used elsewhere in the program.
End Sub
~~~~

~~~~
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
~~~~

~~~~
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
~~~~

After making the major changes to the program's functionality, I added UI elements so that the program was easier to read and use. I started by adding more information to error boxes including a title, tone (audio), and an icon.

~~~~
MessageBox.Show("Enter a numeric value.", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
~~~~

<figure class="single">
    <a href="/assets/images/Inventory Project/Message Box V3.png"><img src="/assets/images/Inventory Project/Message Box V3.png"></a>
</figure>

I also formatted the front UI so it is more open and readable.

<figure class="single">
    <a href="/assets/images/Inventory Project/Project V3.png"><img src="/assets/images/Inventory Project/Project V3.png"></a>
</figure>

The project files can be found below.

<a href="https://github.com/jonathan-rogers-dev/Inventory-Manager/releases/download/V3/honors_inventory_project_v3.exe" class="btn btn--inverse btn--x-large">Executable (.exe)</a>
<a href="/assets/files/Inventory Project/V3/honors_inventory_project_v3.vb" class="btn btn--inverse btn--x-large">Form (.vb)</a>

## Inventory Project Part III Reflection
The Inventory Project was a great method of applying what I had learned in Computer Programing I Honors. I learned how to create, program, design, debug and test a program from scratch utilizing Visual Basic and Visual Studio. Like I said before, I am satisfied with the results of this project, and I look forward to continued development. While making this program, I was able to reinforce my knowledge of using the math class, arrays, procedures, and functions along with other Visual Basic components. I also did have some difficulties in which I learned from. One difficulty I had was using the .Max class to get the largest value from an array. I tried several ideas before I found the solution that I believe works best in this situation. The other difficulty that occurred was when the user would tell the program that they would enter a certain amount of items but only entered a portion of that, the program would display both the items entered and items that were not entered in the list box. This was fixed by using an If statement before the data from the arrays were added to the Listbox. As an additional feature, I also wanted to give the ability to the user to print the data set they created. Unfortunately, I was not able to figure out how to do this and I was not able to release the program with that feature although I do intend to complete more research on this topic and add it to the program.

# More Information
For accessibility and continued development, this project is now available on Github. Please feel free to create issues, fork the repository, or create pull requests. To access the repository, click the button below.

<a href="https://github.com/jonathan-rogers-dev/Inventory-Manager" class="btn btn--inverse btn--x-large">Inventory Project GitHub Repository</a>
