# Unit Assessment: Excel & VBA

You are about to complete your first Unit Assessment! This Unit Assessment allows you to check your knowledge, as well as demonstrate your competency in key concepts from Modules 1 and 2. 

After submitting the assessment here, you will see a summary of your performance. While you will not be able to see your performance on individual questions, you are allowed unlimited attempts to complete the assessment.

Some of the questions on this assessment require specific resources. Download the following resources before you get started. 

[Top5000Songs_Unsolved.xlsx](resources/Top5000Songs_Unsolved.xlsx)

[conditional_formatting.xlsx](resources/conditional_formatting.xlsx)

[ProductionPivot_Unsolved.xlsx](resources/ProductionPivot_Unsolved.xlsx)

[house_data.csv](resources/house_data.csv)

[ePhone_prices.xlsx](resources/ePhone_prices.xlsx)

[star_counter.xlsm](resources/star_counter.xlsm)

**NOTE:** This assessment has questions and answers that are shuffled so they may not be in this order. Correct answers are indicated and bold.

## Question 1

Using `Top5000Songs_Unsolved.xlsx`, what is the mean (to the nearest hundredth) of "final_score" for all 5,000 songs?

- 8.189089
- 8.18
- **8.19 (Correct Answer)** 
- 8.20

## Question 2

Using `Top5000Songs_Unsolved.xlsx`, what is the median (to the nearest hundredth) of "raw_usa" for all 5,000 songs?

- 4.345
- **4.34 (Correct Answer)** 
- 3.31
- 4.74

## Question 3

Using `Top5000Songs_Unsolved.xlsx`, what is the standard deviation (to the nearest hundredth) of "raw_eng" for all 5,000 songs?

- **2.65 (Correct Answer)**
- 2.655
- 2.654
- 2.66

## Question 4

Using `Top5000Songs_Unsolved.xlsx`, what is the median (to the nearest thousandth) "raw_usa" score for songs after 1980?

- 3.51
- 4.34
- 4.340
- **3.508 (Correct Answer)**

## Question 5

Use `Top5000Songs_Unsolved.xlsx`.

1. Select all of the data within the “Top 5000 Songs” worksheet and then create a new pivot table.
2. Make a pivot table that can be filtered by "year" and contains two rows: "artist" and "name. All of an artist's songs should be listed below their name.
3. Update your pivot table to contain values for:
    - how many songs an artist has in the top 5,000
    - the sum of the "final_score" of their songs
4. Sort your pivot table by descending sum of the "final_score."

What is the sum (to the nearest thousandth) of "final_score" for the #3 artist?

- 293.326
- **486.932 (Correct Answer)**
- 610.779
- 486.93

## Question 6

Open `conditional_formatting.xlsx` and select the range A1:A20.

Under “Conditional Formatting,” select “Icon Sets” followed by “Indicators.” Then under icon sets mouse over the indicators section and choose the “3 Symbols (Uncircled).”

What icon is applied to cell A1?

- **Red X (Correct Answer)**
- Exclamation point
- Green checkmark
- Yellow circle
- Green arrow

## Question 7

Open `ProductionPivot_Unsolved.xlsx`.

1. Use a `VLOOKUP()` that references each row's "Product ID" and determine the "Product Price" of each row in the Orders sheet.  The "Product Price" of a row does not include shipping.
2. Use a `VLOOKUP()` that references each row's "Shipping Priority" and determine the "Shipping Price" of each row in the Orders sheet.
3. Next, select all of the data in the Orders sheet and create a pivot table in a new sheet that calculates the sum of both "Product Price" and "Shipping Price" for each "Order Number" and "Product ID."

Using the pivot table, what is the highest total order price (total product price plus total shipping price) to the nearest cent?

- $648.53
- $282.72
- $157.84
- **$320.47 (Correct Answer)**

## Question 8

Using `house_data.csv` and the 1.5*IQR rule, how many house prices are outliers?

Note: Use the `QUARTILE.EXC` function in Excel to calculate quartiles.

- **1159 (Correct Answer)**
- 0
- 566
- $1,129,912.50

## Question 9

Using `house_data.csv` and the rule that outliers are greater than 3 standard deviations from the mean, how many house prices are outliers?

Note: Use the `STDEV.P` function in Excel to calculate standard deviations.

- 13
- **406 (Correct Answer)**
- 21207
- 1159

## Question 10

An analyst sends you a price list of ePhones in a panic. They accidentally deleted one of the prices and didn’t have it stored anywhere else. All they have is a PDF report that says the mean of the prices was $725, and the median price was $750. The other phone prices are $500, $600, $700, $800, and $900. What is the missing price?

- 750
- **850 (Correct Answer)**
- 900
- 1000

## Question 11

That analyst calls you up the next day, in another panic. They have a list of lead concentration measurements from soil samples, but they’ve accidentally deleted another entry. The sample measurements were 35, 45, 45, 50, 70, and 75 ppm. They have a report that says the standard deviation of measurements was exactly 15 ppm. What is the missing measurement? (Hint: use the formula for sample standard deviation)

- 25
- 24.29
- 35
- **65 (Correct Answer)**

## Question 12

A given dataset has a median of 500 and a mean of 765. What can be inferred about the dataset?

- The dataset is skewed to the left.
- **The dataset is skewed to the right. (Correct Answer)**
- The dataset is not skewed.
- Nothing can be inferred from the information given.

## Question 13

Create a new repository on GitHub and initialize the repository with a README. What is the message of the first commit?

- There are no commits yet.
- “README”
- “README.md”
- **“Initial commit” (Correct Answer)**

## Question 14

Open `star_counter.xlsm`.

Create a VBA script that tallies the number of "Full Stars" per student and enters them into the Total column.

What is the average number (to the nearest hundredth) of stars for each student?

- **3.36 (Correct Answer)**
- 3.40
- 3.37
- 3.3

## Question 15

The following code has a syntax error. Which line of code contains the syntax error?


```
1 Sub SumNumbers()

2     Dim total As Integer
    
3     total = 0
    
4     For i = 1 To 100
    
5         total = total + i
        
6     End For
    
7     MsgBox (i)

8 End Sub
```

- Line 1
- Line 2
- Line 3
- Line 4
- Line 5
- **Line 6 (Correct Answer)**
- Line 7
- Line 8

## Question 16

The following code snippet calculates the sum of values in cells A1 to A100, but one line has been obscured with asterisks (***). What is the obscured line? 

```
    *********
    
    For i = 1 To 100
    
        total = total + Cells(i, 1).Value
        
    Next i
```

- `Sub Calculate():`
- **`total = 0` (Correct Answer)**
- `MsgBox("total")`

## Question 17

What keyword is used to initialize a variable in VBA?

- **`Dim` (Correct Answer)**
- `Sub`
- `Next`
- `Run`

## Question 18

What VBA method is used to generate a message box?

- **`MsgBox` (Correct Answer)**
- `MessageBox`
- `Message Box`
- `Msg Box`
- `MgBx`

## Question 19

A junior analyst tells their coworker that the `InputBox` method is used to take input from the user in VBA. Is this a correct statement?

- **Yes, the statement is correct. (Correct Answer)**
- No, the correct method is the `InputUsr` method.
- No, the correct method is the `input()` method.

## Question 20

Open a new Excel file, and write your favorite number in cell A1. Write a VBA macro to format cell A1 as a percentage to a tenth of a percent and set the font style to bold.

Which of the following snippets will run correctly in your macro?

**Option 1: (Correct Answer)**

```
Range(“A1").NumberFormat = “0.0%” 
Range(“A1").Font.FontStyle = “Bold”
```

Option 2:

```
Range(“A1").NumberFormat = vbPercent(1)
Range(“A1").Font.FontStyle = “Bold”
```

Option 3:

```
Range(“A1").NumberFormat = “0.0%”
Range(“A1").Font.FontStyle = vbBold
```

Option 4:

```
Range(“A1").NumberFormat = vbPercent(1)
Range(“A1").Font.FontStyle = vbBold
```
