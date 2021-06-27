# **ANALYZING THE BEST STOCK RETURN OPTIONS**
Kimberly Ellerbe

## **Overview of Stock Analysis Project**
### Purpose
The purpose of this analysis is too evaluate the performance of multiple stocks over 2 sample years, 2017 and 2018, by daily volume and yearly return. The analyst also wants to target a specific stock's performance (DQ) as his relatives are really keen to purchase the stock.  The analyst decided to use Excel to create a macro in Visual Basic that would show which stocks are better choices to invest in by the two measurable criteria. 


### Results
#### Process for Obtaining Results
The process for recommending which stocks performed required a VBA macro that would loop through every row of the data and tally up the volume and determine the yearly return for each stock in the database by hardcoding the database used into the code.  The second process required writing the code a little more "efficiently" to reduce any redundancy in the coding and to avoid hardcoding the database information and instead entering the information requested by a message box prompt.

#### Runtime Data
for 2018 stocks, the refactored script's runtime was discovered to be 28.9% slower than the Mod 2 original script; and, for 2017 stocks, the refactored script runtime was 4.6% slower than the Mod 2 original script code.  It is my assumption that the position of the conditional code statements within the FOR loops to iterate over the rows as well as the dynamic input of the year into the macro via the message box makes a difference in runtime versus when the year is hardcoded within the macro code. (See table below for runtimes)

                    2018            2017          %Diff
                  
  Refactored       1.660156|        1.566406|     28.9%
 
  Original         1.230469|        1.507813|      4.6%


- Add photos of stock analysis & code

![Module3PyPollTerminalResults](https://user-images.githubusercontent.com/79073778/111101194-c6292e00-851f-11eb-8dfb-4ebd3dfcc27a.png)


### Summary In Brief
In this particular code, the advantages of the refactored code simplified the nested loop process as well as made the code more readable.  However, in this project, it could be inferred the disadvantage of the refactored code made the script run a bit slower than the original.  One might generalize that because the code required user input, that may have caused a slow delay in the refactored code's runtime.



![Module3ChallengeVSCode](https://user-images.githubusercontent.com/79073778/111101610-a6463a00-8520-11eb-94a9-df091911fa8c.png)
