# Stock Analysis

## Author
Lydia Delgado Uriarte

## Overview of project

### Purpose
Work with an entire dataset to get to know the stock market over the last few years, using VBA Excel to elaborate a script to manage this dataset making the code efficient and help us with a deep analysis. 

## Results 
### All Stocks 2017
<img width="388" alt="All_Stocks_2017" src="https://user-images.githubusercontent.com/71950779/149392078-f8a6ab04-2a9d-427a-81ef-aa050a8b1f89.png">

### All Stocks 2018
<img width="388" alt="All_Stocks_2018" src="https://user-images.githubusercontent.com/71950779/149394472-ff175451-6b38-42b8-97c6-9eb9bddeb0ef.png">

### Script
#### Original Code
<img width="300" alt="Original_Code" src="https://user-images.githubusercontent.com/71950779/149398411-e8fa827c-d435-4a5b-b5f2-55f62cc0e004.png">
<img width="404" alt="Original_Code_Pt2" src="https://user-images.githubusercontent.com/71950779/149398420-10e98417-c6df-477a-9ce2-1195a383832a.png">
<img width="314" alt="VBA_First_Time_2017" src="https://user-images.githubusercontent.com/71950779/149404527-7f8b3df6-634e-42ab-9af3-727fa229dce9.png"><img width="319" alt="VBA_First_Time_2018" src="https://user-images.githubusercontent.com/71950779/149404567-328c4206-e3d4-4fcf-8d62-fd0a9c18d2e6.png">


#### Code Refactored
<img width="541" alt="Refactored_Code" src="https://user-images.githubusercontent.com/71950779/149398428-95c018aa-6663-4a40-bfd0-eede8f89cad7.png">
<img width="320" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/71950779/149301729-61b08e85-3780-4b09-b0ce-e841639708d6.png"><img width="314" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/71950779/149392019-82342345-5d87-452e-8ec8-6cf28585a60f.png">


### Analysis
Seeing this from perspective, it's understandable how DQ could be a great company to invest according to 2017 stocks, however the stocks in 2018 show otherwise being **DQ** the ***worst ticker*** to invest in with **-62.6%** of stocks. The ***best option*** to invest in is **RUN**, with the results presented above we can see that the stocks represent a **84.0%** and is by far the best option among the other companies presented.

The execution in 2017 stocks and 2018 stocks time's are very close related to how fast the code compiles, the one with best performance is the one of 2018. If we compare the performance of the refactor code and the original code, the refactor code starts working immediately and the original code takes it time to start compiling. 


## Summary
### Refactoring code in general
#### Advantages
A well performance in how long it takes the code to run can be due the number of lines it has and how simplified it is, the original code possibly could be a mess so by refactoring the code we organize this mess and make it understandable for the reader how it works.

#### Disadvantages
Refactor code could bring some problems, since you can miss out important information trying to simplify the code. Is important to keep in mind that the point of refactoring is improve the performance by keeping useful information not just deleting lines.

### Refactoring the original code and refactored VBA script
#### Advantages
In this case, refactoring the code is organized in lists and there is no need to use another for loop for the volume since in one we can keep all our information about the volume and return, also it simplifies the time performance making the program rust faster and is cleaner.

#### Disadvantages
Refactoring this code was confusing at the time of creating new arrays to store the information, considering this information was not stored by ticker, but also that when testing this it could bring some errors and it did not work at the first try.
