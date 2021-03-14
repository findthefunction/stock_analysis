
# Stock analysis exercise with VBA

## 1. Project Overview

Preparing a simple interface to enable the rapid review of historical stock data, to be formulated in an efficient manner that is flexible for future flexibility.  

The project starts with a simple analysis of twelve different stocks over a time period of two years.  Once the code is written and functioning the goal is to observe how it functions and find methods to improve its performance.  The code is refactored to run more efficiently with the intent on it being able to handle larger volumes of stock data more efficiently.  

This is accomplished by creating several output arrays and using the loop function effectively to maximize data gather and minimize the number of times the information must be cycled to obtain the desired results.


## 2. Results

By refactoring the original code as written we were able to see relatively significant improvement.  The refactored code saw a ~390 percent increase in speed.  The original code took 1.25 second to compile the information. 

![original_stock_analysis_time_cropped](https://user-images.githubusercontent.com/31022640/111081378-4324bb00-84c0-11eb-9b43-a46b4972d783.png)!

The improved subroutine blew its doors off with a time of 0.32 seconds.  While a second or two might not seen like a long time for humans, when running large data sets it could amount to significant resource expenditure and processor load. 

![green_stocks_optimized_cropped](https://user-images.githubusercontent.com/31022640/111081411-75361d00-84c0-11eb-93ed-6af3e10357a5.png)

Multi-core threading cannot be done natively in VBA which significantly hinders its ability when it comes to large sets of data and numerous operations. 

## 3. Summary

Running efficient code is a tenet in the programming world.  Writing efficient and effective code is what one should strive for.  Less moving parts equals less parts to break.  The advantage of this exercise was to display an iteration of code that functioned according to the intended goal, and then to analyse and see very clearly the speed advantage of the refactored subroutine. 

There are a few instances where grouping subroutines to gain efficiency might have disadvantages.  When working on larger project with several people working together there could be conflicts if different components need to extract data from different point in the loop. You can also isolate variables for reuse in other parts of code which may cause an efficiency disadvantage but overall, the code could be more reusable. 

For this instance, running this VBA subroutine on a modest data set there are not very many disadvantages. By using Dim statements, we are able to tell VBA about a variable we will use later which make the code more agile and reusable.

VBA is a great tool for convenient automation within the Excel environment.  







