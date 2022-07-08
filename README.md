# Stock Analysis using Excel & VBA

## Overview of Project

The purpose of this project is to use stock market data to learn Visual Basic for Applications (VBA) with Excel. We will use VBA to analyze and make calculations for several companies stock market performance. The criteria was simply to measure each company's total volume and return for the years being analyzed.The completed analysis can then be used by the end user to determine if one company would be a better investment than the others. 

The VBA script that runs the analysis was written with the existing data in mind from just two years of information for 12 companies. The VBA script appeared to run efficiently. However, with the concern that it could be used on a much larger dataset we refactored the code to make it run faster and more efficiently.

## Results of Analysis

### Stock Performance Comparison Between 2017 & 2018

As you can see below, for this group of green stocks 2017 was a good year as all but one of them had a positive return for the year. Tickers DQ had the best performance of the bunch followed by SEDG. As a group, green stocks appear to have performed well for the year. 

![2017 Analysis Results](/Resources/2017_Analysis_Results.png)

While 2017, was a great year for this batch of green stocks 2018 looks to have been a more challenging year. All but two of the company's public stock had a negative return for the year. Only EPNH and RUN provided a postive return for investors for the year.

![2018 Analysis Results](/Resources/2018_Analysis_Results.png)



### VBA Execution Times: Original Script Compared to Refactored Script

In this section, we will look at and discuss the perfomance of the original VBA script to the refactored VBA script. SHown below, the original version of the VBA script ran the analysis in approximately 0.45 to 0.46 seconds.

![2017 Original Code Run Time](/Resources/VBA_Challenge_2017_Original_Code.png)
![2018 Original Code Run Time](/Resources/VBA_Challenge_2018_Original_Code.png)

While the refactored VBA script consistently completed the analysis in 0.625 seconds.

![2017 Refactored Code Run Time](/Resources/VBA_Challenge_2017.png)
![2018 Refactored Code Run Time](/Resources/VBA_Challenge_2018.png)

As a result of these measurements, the refactored code is multiple times faster than the original code.


## Summary
Summary statement that addresses two questions
1. What are the advantages & disadvantages of refactoring code.
2. How do these pros & cons apply to refactoring the original VBA script.


