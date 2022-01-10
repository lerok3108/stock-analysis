# stock-analysis
## Overview of Project
In this project I worked with a dataset that contained information about stocks for 2 years, 2017 and 2018 stored in Excel spreadsheets. The set contained information about stocks' price (open, high, low, close) and volume traded. I created a code using VBA that analyzed this information and presented the stock performance results for each year in an easy-to-read form. 
### Purpose
The purpose of this project was to refactor the code I created during the class modules to make it run more efficiently and potentially be applied to the whole market and not just 12 stocks utilized in this project. 

## Results

Refactoring the initial code included the process of eliminating some steps by assigning variables/creating arrays and thus getting rid of nested loops. The refactored code ran a second faster than the initial version of the same code due to improved efficiency. Additionally, the refactored version of the code included the formatting module which had to be ran separate in the initial version thus saving even more time. 
### Results of the refactored code


![2017_VBA_challenge](https://user-images.githubusercontent.com/96098938/148711256-3bffcf83-384e-4794-be3b-69aa08305504.PNG)
![2018_VBA_Challenge](https://user-images.githubusercontent.com/96098938/148711262-f37bb072-64af-4483-8851-e6fea1f32685.PNG)

### Results of the initial version of code

![2017 All stock analysis](https://user-images.githubusercontent.com/96098938/148711294-79783b34-4112-4e62-a8b2-c2a7762c9848.PNG)
![2018 All stock](https://user-images.githubusercontent.com/96098938/148711301-f2a40edd-4db6-4215-b2d1-db8c23acad3e.PNG)

### Refactored Code
![Refactored Code](https://user-images.githubusercontent.com/96098938/148711499-22e29a7a-638a-4e9d-b151-89e07bb89ca1.PNG)


### Initial Code

![All Stock Analysis Initial Version](https://user-images.githubusercontent.com/96098938/148711506-d7702ab1-8170-4bf3-924a-c3533f90ebdd.PNG)


## Advantages
The advantages of the refactored code include 
-Improved efficiency and logic, faster run time
-Less redundance
-Possibility of expanding the code to be able to analyze more stocks than initial 12  in the current example
-The code is easier to read/understand by other users

## Disadvantages

-Refactoring the code that already works is always risky
-If the code gets bigger, refactoring process will take longer
