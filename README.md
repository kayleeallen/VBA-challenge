# VBA-challenge

## Performing analysis using VBA in Excel to analyze stock trends

### Overview
The purpose of this challenge is to refactor code to make it run more efficiently. 

### Background
For this project, we received a dataset regarding information of various stocks for year 2017 and 2018. Using VBA, we had to create a worksheet to analyze the stocks for both year 2017 and year 2018. We received a piece of code that we had to refactor so it could run more efficiently. We tested the efficiency of the code by calculating the run time of the code.

### Original Code Runtime

Using the original code to run an analysis of the stocks for year 2017, the run time was as shown:

![VBA_Challenge_2017_Original](https://user-images.githubusercontent.com/90978520/135771546-5462e5a5-5fbd-4fe2-a554-6e4e6522ae52.png)


Using the original code to run an analysis of the stocks for year 2018, the run time was as shown:!

[VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/90978520/135771554-a5c291a4-9571-4f83-af7b-135ce6b5c9fa.png)

### Refactoring the Code

#### 1. Set the ticker index to 0
   
   ![Set_tickerIndex_0](https://user-images.githubusercontent.com/90978520/135772489-588ebf0d-f751-4e68-8a42-eaa88cc8110a.png)


#### 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices 

  ![Arrays](https://user-images.githubusercontent.com/90978520/135772511-8b3d1bcc-662e-4a8e-a3dd-58fa79008d37.png)
  
#### 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays

  ![README_3](https://user-images.githubusercontent.com/90978520/135772656-86de7ab8-1dc3-484a-887c-2849b3ba98c4.png)

#### 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices

  ![startingPrices_endingPrices](https://user-images.githubusercontent.com/90978520/135772730-7512c96f-2213-4ff2-8e2a-4e4ea86bc6fe.png)

#### 5. Code for formatting the cells in the spreadsheet is working

  ![Formatting_Code](https://user-images.githubusercontent.com/90978520/135772764-aa2c11c5-2460-4f4d-b892-b337ded06dbb.png)

#### 6. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

  Outputs for 2017 Analysis:
  
  ![image](https://user-images.githubusercontent.com/90978520/135772805-b5752a5d-6f83-406b-b6f2-b978a604cf4d.png)
  
  Outputs for 2018 Analysis:
  
  ![2018_Analysis](https://user-images.githubusercontent.com/90978520/135772818-0ed9a9af-a279-4319-be90-b3bc360a9f78.png)


  
