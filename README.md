# stock-analysis-Analyzing stocks with VBA

## Overview of Project

  The purpose of this analyis was to help Steve figure out which stocks his parents should invest in.
His goal is to analyze a spreadsheet with a number of different given stocks, labeled with their 
respective volumes of **stocks traded**, their **opening dates**, and their **highest** and **lowest** values traded.
I wrote different VBA scripts to automate calculations for each stock's **starting prices**, their **ending prices**,
and their **total daily volumes**, as well as each of their **return values**. This data workbook, along with their
VBA *macros*, aimed to assist Steve in deciding which stocks would be best for his parents to invest in.

## Results

We compared stock prices between two years: 2017 and 2018.

### 2017 and 2018 stock comparison

Here are the results from 2017's analysis:

![2017_Analysis](https://github.com/Kyle2Miles93/stock-analysis/blob/main/2017%20All%20Stocks%20Analysis.png)

and here are the results from 2018's:

![2018 Analysis](https://github.com/Kyle2Miles93/stock-analysis/blob/main/2018%20All%20Stocks%20Analysis.png)

  As is shown, 2017 was a very good year to invest in almost every one of these 12 companies, except for the stock corresponding to the 
ticker: **TERP**, which had a negative return of 7.2%. There were four of these twelve stocks that had increases of over 100% for their return percentages!

  In contrast, 2018 was a very bad year to invest. all but two of those same twelve stocks decreased in value overall in 2018. The only two that increased were
the stocks with corresponding tickers: **RUN** and **ENPH**, which increased by 84.0% and 84.9%, respectively. In addition, almost every ticker's total volume in 2017 with some exceptions were lower thaqn 2018's, but 2017's return values were all better than the ones in 2018! I should conduct further analysis to 
find out why that was the case for those years.

### First VBA code vs Refactored VBA code

Here is a screenshot of the run time of the refactored code for 2017. It is about four-tenths of a second faster than the first version's code.
The messagebox shows the runtime for this program:

![Refactored VBA Analysis 2017](https://github.com/Kyle2Miles93/stock-analysis/blob/main/VBA_Challenge_2017.png)

Here is 2018's runtime:

![Refactored VBA Analysis 2018](https://github.com/Kyle2Miles93/stock-analysis/blob/main/VBA_Challenge_2018.png)

Based on the difference of execution times for each script, I conclude that if there were, instead of twelve different stocks to analyze, say, *one thousand*,
then the time saved as well as the costs saved from running the programs would be very significant. This is because the run times were nearly four times faster 
in the refactored code than the original code.

## Summary

I conclude that the advantges of refactoring code are:

1) **faster runtime for faster work completion**
2) **money is saved for the company because of less computing power used per run**
3) **less chance for the computer to crash because of segmentation faults and less chance of errors**

The disadvantages may include:

1) **The refactored code was challenging to figure out how to make work,**
**though a more experienced programmer would probably have no problem.**
2) **The refactored code uses arrays instead of one-value variables for**
**the ticker, volumes, and starting/ending prices. This would**
**make it more finicky if the programmer was inexperienced.**

![Original Code](https://github.com/Kyle2Miles93/stock-analysis/blob/main/Original%20Code.png)

![Original Code Pt. 2](https://github.com/Kyle2Miles93/stock-analysis/blob/main/Original%20code%202nd%20part.png)

![Refactored code](https://github.com/Kyle2Miles93/stock-analysis/blob/main/Refactored%20Code.png)

![Refactored Code Pt. 2](https://github.com/Kyle2Miles93/stock-analysis/blob/main/Refactored%20Code%202nd%20Part.png)










  
