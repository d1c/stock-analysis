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

From the refactoring exercise completed and discussed above in this case there was a clear advantage to refactoring the code. If the user decides to use the VBA script on a much larger data set then the efficiency of the refactored code will definitely be pronounced. 

Doing some research on the Internet there are many advantages and disadvantages to refactoring code:
- Advantages:
  - Efficiency: A goal of refactoring code is often to improve the efficiency of the code. For example, in our case the refactored code can do the same "job" faster than the original code.
  - Maintainability: Refactored code should be easier to read and maintain. If the code will be enhanced with new features in the future then refactoring could reduce the chances of future bugs by making the code easier to use or reuse.
  - Removing Bad Code: When schedules are tight, code is written in the best way possible in the time given. Refactoring can catch these issues and rewrite the code to be more efficient.

- Disadvantages
  - Expensive: Time spent refactoring can take away from valuable time spend developing new features.
  - New bugs: When touching code there is always the possibility of introducing new problems.
  - Tight Delivery schedule: In past jobs where we put a lot of pressure on vendors to deliver new features time spent refactoring would have caused delays in providing users the features they needed.
- Source: [Pros And Cons Of Code Refactoring](https://www.c-sharpcorner.com/article/pros-and-cons-of-code-refactoring/)

3. How do these pros & cons apply to refactoring the original VBA script.

Refactoring the VBA code in my case became a learning experience; However, resulted in much faster code. The learning experience resulted from having to rethink through the problem the code was written to solve to find a different solution that ended up being faster. I can see a situation where the goal is faster code but the different solution could end up being worse.

From a maintainability standpoint, refactoring was another opportunity to add additional comments to the code making it more readable. Improved readability will be important in the future should the code need to be reused, changed or adapted.

The learning experience came when I fat fingered a couple of statements. I then had to go back through the code to figure out why the VBA script was not running successfully. It made me realize why I have often heard software developers express reluctance to touch working code.

