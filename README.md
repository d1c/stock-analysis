# Stock Analysis using Excel & VBA

## Overview of Project

The purpose of this project is to use stock market data to learn Visual Basic for Applications (VBA) with Excel. VBA is used to analyze and make calculations for several companies stock market performance. The criteria was simply to measure each company's total volume and return for the years being analyzed. The completed analysis can then be used by the user to determine if one company would be a better investment than the others. 

The VBA script that runs the analysis was written with the existing data in mind from just two years (2017 & 2018) of information for 12 companies. The VBA script appeared to run efficiently. However, knowing the VBA script could be used on a much larger dataset we refactored the code to make it run faster and more efficiently.

## Results of Analysis

### Stock Performance Comparison Between 2017 & 2018

As you can see below, for this group of green stocks, 2017 was a good year. All but one of companies had a positive return for the year. Ticker DQ had the best performance of the bunch followed by SEDG. As a group, green stocks appear to have performed well for the year. 

![2017 Analysis Results](/Resources/2017_Analysis_Results.png)

While 2017, was a great year for this batch of green stocks, 2018 was a more challenging year. All but two of the company's had a negative return for the year. Only EPNH and RUN provided a postive return for investors for the year.

![2018 Analysis Results](/Resources/2018_Analysis_Results.png)

Using just two years of data it is difficult to come to a conclusion if one of the 12 companies would be the best investment going forward. This analysis is a good start; however, further analysis would be needed to come to a decision.

### VBA Execution Times: Original Script Compared to Refactored Script

In this section, we will look at and discuss the perfomance of the original VBA script to the refactored VBA script. Shown below, the original version of the VBA script ran the analysis in approximately 0.45 to 0.46 seconds.

![2017 Original Code Run Time](/Resources/VBA_Challenge_2017_Original_Code.png)
![2018 Original Code Run Time](/Resources/VBA_Challenge_2018_Original_Code.png)

While the refactored VBA script consistently completed the analysis in 0.625 seconds.

![2017 Refactored Code Run Time](/Resources/VBA_Challenge_2017.png)
![2018 Refactored Code Run Time](/Resources/VBA_Challenge_2018.png)

As a result of these measurements, the refactored code is multiple times faster than the original code. While the original code was fast enough for the dataset being analyzed for this project the refactored code will allow fast performance if the user expands the data being analyzed to a much larger dataset.

## Summary - Refactoring: Advantages and Disadvantages

### Advantages
Some advantages of refactoring code include improved efficiency, maintability and removal of "bad code." As data sets grow code that runs efficiently becomes more important. For example, data has to be processed and in many cases ingested into databases prior to being analyzed. The larger the volume of data the bigger impact more efficient code becomes. In one case I am familiar with, refactoring code saved almost two hours to process daily data making the data available for analysis much sooner.

### Disadvantages
Refactoring code clearly has its benefits but they do come with costs. Some disadvantages are increased expense, chance of introducing new bugs, and impacting tight delivery schedules. Refactoring takes time so therefore increases the cost and delivery of software. When working code is touched the risk of introducing new bugs is also increased. 

### Advantages and Disadvantages of Scripts
It could be said that a disadvantage of the original VBA script is that it successfully did what it was designed to do. The original script is clearing at a disadvantage when it comes to execution times. If this is the only dataset the original VBA script will be used to analyze then the added time to refactor the code could be considered wasted since it saved less than half a second of time. It is probably not a safe assumption the the script would never be used with much larger data sets. The use case of this project was to analyze the performance of 12 companies stock performance. If the dataset were expanded to every stock listed on American stock exchanges then the faster execution of the refactored VBA script would be a much bigger advantage.





In my case, while refactoring the VBA code resulted in faster code the process became a learning experience. The learning experience came as a result of having to think through and already solved problem to find a different more efficient way to solve the problem with VBA. While working through the refactoring exercise it occurred to me that while the goal is better or faster code the end result could be the opposite if not written carefully.

From a maintainability standpoint, refactoring was another opportunity to add additional comments to the code making it more readable. Improved readability will be important in the future should the code need to be reused, changed or adapted.

Another learning experience came when I mistyped a couple of statements. I then had to go back through the code to figure out why the VBA script was not running successfully. The process of finding the mistyped code was time consuming and even a little time consuming. It made me realize why I have often heard software developers express reluctance to touch working code.

