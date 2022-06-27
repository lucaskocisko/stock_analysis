#Stock Analysis using VBA

##Overview
Steves parents are looking to invest in GE (green energy) stock. After reviewing the requested stock, we then made a new macro comparing this option to others. After writing various VBA macros exploring lists, for loops, if/else statements, etc. I made a a script to build a table for the input year. This submission is a refactored version of that macro.

##Results
###2017
GE stock performed on average very well. 4 different options returned over 100%
![](https://github.com/lucaskocisko/stock_analysis/blob/main/2017unfactored.png)
![](https://github.com/lucaskocisko/stock_analysis/blob/main/VBA_Challenge_2017.png)
![](https://github.com/lucaskocisko/stock_analysis/blob/main/2017table.png)
At first, I set different variables of the ticker index as the wrong data type and table was slightly off. After refactoring I saw my problem (dim starting&ending prices as single).

###2018
GE stock performed worse the following year. Only two options ended in the green, both at low 80%.
![](https://github.com/lucaskocisko/stock_analysis/blob/main/2018unfactored.png)
![](https://github.com/lucaskocisko/stock_analysis/blob/main/VBA_Challenge_2018.png)
![](https://github.com/lucaskocisko/stock_analysis/blob/main/2018table.png)

##Summary
pros and cons of refactoring

The main advantage to refactoring code is the improved flow of your code. After already seeing the code from start to finish you have a good concept of it. By trimming and notating your work its more presentable to someone whos never seen your problem. My finished code was more accurate and ran faster.

A disadvantage to refactoring is the amount of time it takes. It's also important to not be careless and create new bugs.
