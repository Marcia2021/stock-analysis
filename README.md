# Stock-Analysis Using VBA

## Overview of Project

### Purpose

During the Module practice, we have used VBA to help Steve to create a workbook by analyzing green stock data. This analysis was focus on evaluating the return, the difference between the ending prices and the starting prices, in the analysis year of selection for every stock in the dataset. During the analysis, we drafted a VBA script used nested For Loop to calculate the values, which met our initial purpose. However, the efficiency for a VBA script is highly depends on how to process the data. We are going to refactor the VBA script to speed up the process.  

## Analysis

The original VBA script used nested For Loop to loop through the data. The first For Loop was used to loop through the 12 tickers, then the nested For Loop was used to loop through all the rows in the green stock data to find matched records to calculate each value based on the criteria. The original VBA script also included the output step inside the loop, which means the VBA needed to reference to the worksheet 12 times.  

When refactoring the VBA script, we are going to use one layer of the For Loop to loop through all the rows, and create the values for each variable at the same time. Then use anther For loop to output all the values for each variable to the worksheet at once. 

1.	As in the initial VBA script, created the “startTime” and “endTime” variables use “Dim”. Both of these variables will be “Single”.

    ![pos1](https://user-images.githubusercontent.com/79289806/109896432-eb37ba00-7c5e-11eb-986d-76b0afc20782.png)
 

2.	Used InputBox function to display a prompt in a dialog box to input the year for analysis. The year for analysis will be store in the variable “yearValue”. This variable will be used in the following steps.
 
    ![pos2](https://user-images.githubusercontent.com/79289806/109896433-ebd05080-7c5e-11eb-844c-1530da2fc31c.png)
 
    The dialog box:
 
    ![pos3](https://user-images.githubusercontent.com/79289806/109896434-ebd05080-7c5e-11eb-93c1-ef605034bbd4.png)
    
    The green stock data only contains 2017 and 2018 data. We could enter 2017 or 2018 in the dialog box to analyze the data for each year respectively.

3.	After the initial setting, created a table frame in a new worksheet called “All Stocks Analysis”. Activated the worksheet, used Range and Cells to create the table frame. 

    ![pos4](https://user-images.githubusercontent.com/79289806/109896435-ebd05080-7c5e-11eb-8846-a8382e2be957.png)
 

4.	In the new worksheet, used “Dim” to create an array that carries all the tickers in the 2017/2018 data. There are 12 tickers in the data, used tickers(11) to create 12 associated values. In this step assigned values to each ticker in the array. The tickers array will be used to create the rows in the new table. 
 
    ![pos5](https://user-images.githubusercontent.com/79289806/109896436-ebd05080-7c5e-11eb-8d31-1f9478a1438f.png)

5.	Activated the worksheet for the analysis year.

    ![pos6](https://user-images.githubusercontent.com/79289806/109896438-ebd05080-7c5e-11eb-998c-12f758466b9b.png)

6.	In the activated worksheet, calculated the total number of rows, and store the value in variable “RowCount”.

    ![pos7](https://user-images.githubusercontent.com/79289806/109896440-ebd05080-7c5e-11eb-8384-4dd51cd452a7.png)
 
7.	In the original VBA script, the next step created the first For Loop to loop through all the tickers. Then created a nested For Loop to loop through all the rows. In the refactored VBA script, initialized an index variable “tickerIndex” to 0, which will be used to create outcome variables for each ticker. Additionally, created three output arrays which store the values for each ticker. The number of the values in each outcome array is the same as the tickers array.

    ![pos8](https://user-images.githubusercontent.com/79289806/109896441-ebd05080-7c5e-11eb-9dca-b588b676ea69.png)
 
8.	Used For loop to loop through all the values in the array, used “tickerIndex” to initialize the tickerVolumes(tickerIndex) to zero. This variable will be further processed in the next For loop to cumulate the value of Volume from each year. 
 
    ![pos9](https://user-images.githubusercontent.com/79289806/109896443-ec68e700-7c5e-11eb-80e3-15193eb33ba4.png)
    
9.	Used separate For loop to loop through all the rows in the analysis year of data. Created each value in the tickerVolumes, tickerStartingPrices and tickerEndingPrices arrays. When creating the tickerEndingPrices, if it reaches to the last row for each ticker, increment the tickerIndex. 

    ![pos10](https://user-images.githubusercontent.com/79289806/109896444-ec68e700-7c5e-11eb-897c-8219e62a1e36.png)

10.	After the calculation, activated the “All Stocks Analysis” worksheet. Created a new For loop to output the values of each variable to the table. Then formatted the table. 

    ![pos11](https://user-images.githubusercontent.com/79289806/109896446-ec68e700-7c5e-11eb-88a9-1d81920256b5.png)

11.	At beginning of the script, created startTime

    ![pos12](https://user-images.githubusercontent.com/79289806/109896447-ec68e700-7c5e-11eb-9ead-0780d32f7346.png)
 
    At the end of the script created endTime
    

    ![pos13](https://user-images.githubusercontent.com/79289806/109896448-ec68e700-7c5e-11eb-8727-944e2a600448.png)
    
 
    Calculated the difference between the startTime and endTime, used MsgBox to show the running time.
    

    ![pos14](https://user-images.githubusercontent.com/79289806/109896449-ec68e700-7c5e-11eb-902a-441e104161aa.png)

## Results

From the refactored VBA script, we were able to create tables for 2017 and 2018 respectively.

2017:

![pos15](https://user-images.githubusercontent.com/79289806/109896450-ec68e700-7c5e-11eb-8f38-2ff5ebe84508.png)
 
2018:
 
![pos16](https://user-images.githubusercontent.com/79289806/109896451-ec68e700-7c5e-11eb-857d-ab92f7529ea0.png)

In the meantime, we were able to collect the running time for each year:

2017:
 
![pos17](https://user-images.githubusercontent.com/79289806/109896453-ed017d80-7c5e-11eb-8fd4-269e1d9def7b.png)

2018:
 
![pos18](https://user-images.githubusercontent.com/79289806/109896430-eb37ba00-7c5e-11eb-8a14-60d9d23fdd71.png)

## Summary – Pros and Cons

The initial and refactored VBA script are very similar. The major difference is how to loop though the data. 

In the initial script used the nested For loop which has the loop for tickers as the first layer, then the loop for rows of the data as the second layer. This means when creating the outcome variables for each ticker, the code will loop through the whole data multiple times. Additionally, the initial script included the output step inside the same loop, which requested to reference to the worksheet multiple times. This set up will slow down the process.

In the refactored script, we created an index variable and arrays to store the values for each outcome variable before loop through the data. Then used the index variable by incrementing the value to the largest number of the array to loop though the data once to create all the outcome variables. Additionally, created a separate For loop to output the outcomes to the final table in the worksheet, which requested to reference to the worksheet only one time. 

Here is the comparison of the running time between the initial script and refactored script:

![pos19](https://user-images.githubusercontent.com/79289806/109896431-eb37ba00-7c5e-11eb-9278-57a4ab4e516a.png)

From the table above, it is obvious that the refactored script runs faster than the initial script. 

There are lots of factors that will affect the efficiency of a program.  The green stock data only contains little over 3,000 rows of records. When dealing with larger data, efficiency of the program becomes important. 

1.	The complexity of the program:

    According to the purpose of an analysis and the structure of the data single layer loop and nested loop will be performed differently. For the green stock data analysis,         compared to nested For Loop, single layer of loop will reduce the time of execution. In the meantime, when it is applicable separate loops will simplify the program and         improve the efficiency of execution.

2.	Size and structure of the data:

    The size of data will be another factor that affect the choice of the way of programming. The green stock data is small and requested to create very limited variables. Both     initial and refactored script could complete within 1 second. When processing large size of data, we need to consider the structure of the data and the purpose of the           analysis, nested loops might be necessary for processing the data in sections. Under this circumstance, simple loop might not be a better choice. 
