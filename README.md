#Refactoring the VBA script to improve the performance by measuring the process time 

##Refactoring the VBA code successfully improves the performance speed by approximately 80% faster than original script. 

The original script, without formatting worksheet, takes almost 1 second to complete the task in both 2017 and 2018 (see result [Original VBA 2017](https://github.com/Yunaka1269/stock-analysis/blob/master/Resources/VBA_Original_2017.png) and [Original VBA 2018](https://github.com/Yunaka1269/stock-analysis/blob/master/Resources/VBA_Original_2018.png)). The refactored script reduces the processing time down to around 0.2 second (80% faster) to complete the task including the formatting worksheet in both 2017 and 2018 (see result [Refactored VBA 2017](https://github.com/Yunaka1269/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png) and [Refactored VBA 2018](https://github.com/Yunaka1269/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)). I believe the main driver that helps to improve the speed is separation of nested ***for*** loop. Using the array of tickerIndex enables one ***for*** loop to calculate the totalVolumes without identifying the start/end rows for each ticker and another independent ***for*** loop to look up the startingPrices and endingPrices based on Indexed ticker name. With that said, refactoring saves Excel going through more than 30,000 rows.   

###Summary
1. Advantages and disadvantages of refactoring code 
    -Advantage of refactoring code is to restructure the patterns so that it would eliminate the extra steps and require less memory to run the procedure. Not only improving the  performance, refactoring may help clearing up the script to look easily readable. It also may detect/encounter the bugs and duplicate during the refacotring process.     
    -Diadvantage of refactoring is that it may become cost/time consuming. The first attempt won't necessarily be the the best way to enhance the program. Refactored script may become more complex than the original.  

2. Pros and cons listed above applies to refactoring this challenge. But I personality believe gaining .8 second would not justfy the amount of time spent for refactoring.
Of course, if the dataset was much larger and required to run everyday, then I would save so much time in the long run. before refactoring the script, the evaluation must be done. 
