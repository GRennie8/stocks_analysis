# VBA of Wall Street

## Overview of Project

For this challenge I used the Module 2 workbook code I prepared for Steve to get a better picture of the stock market. The original code used a nested for loop to go through the dataset and display the "Total Daily Volume" and "Return" for 12 stocks. In this project I was required to refactor the code to make it more efficient and to run the script faster.



## Results

### Difference in Code

The original code had a nested for loop which resulted in a significant number of iterations. Shown below is the original code.

<img width="1094" alt="Screen Shot 2022-05-15 at 7 46 16 PM" src="https://user-images.githubusercontent.com/104115586/168508460-2c04abb7-27cb-48f1-9119-40109524d3b2.png">

After refactoring the code, there is now only one loop and the output is in a separate loop. Meaning there is much less work for the code to do.

<img width="1076" alt="Screen Shot 2022-05-15 at 7 45 16 PM" src="https://user-images.githubusercontent.com/104115586/168509153-db67b944-459d-4ec2-b51f-57c614c27197.png">




### Run Time

#### Run time for 2017 with original code

<img width="406" alt="Screen Shot 2022-05-15 at 7 59 04 PM" src="https://user-images.githubusercontent.com/104115586/168508687-3e81072a-e2e0-4ee0-b03b-3e5705d90f97.png">

#### 2018

<img width="405" alt="Screen Shot 2022-05-15 at 7 59 28 PM" src="https://user-images.githubusercontent.com/104115586/168508764-f1667942-64f4-4bec-b240-84c07b4c6086.png">

Both scripts run in about 0.25 seconds.

#### Refactored Code Run Time for 2017

<img width="408" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/104115586/168509431-b8d5c0d2-3330-4fcf-a27a-23a0782eb572.png">


#### 2018

<img width="408" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/104115586/168509435-81814514-56d3-4c2a-b734-86f97fc4c379.png">

It's clear that after the refactoring process, the script runs much faster. Around 5 times faster!

## Summary

#### Advantages and Disadvantages of Refactoring Code

The main advantage of refactoring (if done right) is that the code will run faster or more streamlined due to some unnecessary steps being taken out. Another advantage is if a colleague carries out the refactoring they can do so with a fresh perspective which could lead to finding repeated code or mistakes.

The second advantage could also be seen as a disadvantage due to the risk that a colleague starts to refactor without knowing the original creators intentions could lead to confusion. It can also be relatively easy to get lost in the code while trying to copy or add pieces of code into or out of the original.

#### How Do These Pros and Cons Apply to Refactoring This Script

After the refactoring of the module code it ran significantly faster and was easier to follow in the VBA editor. 
