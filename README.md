# An Analysis of All Stocks
> Performed data analysis to uncover daily volume and yearly return for multiple stocks.
Click here to view the Excel file: [VBA Challenge - Stock Analysis](https://github.com/gforce2332/stock-analysis/blob/master/VBA_Challenge.xlsm)


## Table of Contents
* [Overview of Project](#overview-of-project)
* [Results](#results)
* [Screenshots](#screenshots)
* [Summary](#summary)

## Overview of project
- Using VBA code that interacts with excel to gain valuable insights into the performance of twelve (12) different stocks.
- Collect individual stock information in the year 2017 and 2018 and determine which stocks are worth investing in.
- Calculate daily and yearly volume to get a better idea of how often each stock gets traded.
- Analysis will create insights into which stocks have a greater return and are thus a more valuable stock to invest in.


## Results
Since Daqo dropped over 63% in 2018 it's most likely not the best stock to invest in.
By analyzing multiple stocks better choices can be found. By comparing the stock performance between 2017 and 2018 we can see that stocks dipped quite a bit in 2018 but two stocks,
ENPH and RUN still had a positive rate of return in 2018.

In order to run analyses on all of the stocks, a program flow was created to loop through all of the stock's ticker prices. By writing statements and assigning current starting and ending prices, total daily volume and rate of return can be calculated.

RowCount code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists 
rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
 
'![All Stock Analysis Code](https://user-images.githubusercontent.com/98711219/161366912-be3a69ea-59dc-4a4d-8f37-5e275b03839d.png)




## Screenshots
![VBA_Challenge_2017](https://user-images.githubusercontent.com/98711219/161366922-6e06ccce-f88b-40a2-858d-68e37e8e3df2.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/98711219/161366925-7d5cd687-08a0-4666-8466-926b48852e49.png)


## Summary
* The biggest advantage of refactoring is that it leads to better quality code, makes it easier to understand and faster programming.
* A potential disadvantage of refactoring is that it could be risky when the application is big or the code is long.
* Refactoring made the VBA script run faster. Both the 2017 and 2018 analyses ran originally in 0.5 seconds. After refactoring the 2017 analysis ran in 0.094 seconds and 2018 in 0.078 seconds
  as seen in the screen shots above. 
 




