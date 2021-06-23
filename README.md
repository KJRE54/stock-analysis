# **ANALYZING THE BEST STOCK RETURN OPTIONS**
Kimberly Ellerbe

## **Overview of Analysis Project**
### Purpose
The purpose of this analysis is too evaluate the performance of multiple stocks over 2 sample years, 2017 and 2018, by daily volume and yearly return. The analyst also wants to target a specific stock's performance (DQ) as his relatives are really keen to purchase the stock.  The analyst decided to use Excel to create a macro in Visual Basic that would show which stocks are better choices to invest in by the two measurable criteria. 


### Results
#### Process for Obtaining Results
The process for recommending which stocks performed required a VBA macro that would loop through every row of the data and tally up the volume and determine the yearly return for each stock in the database by hardcoding the database used into the code.  The second process required writing the code a little more "efficiently" to reduce any redundancy in the coding and to avoid hardcoding the database information and instead entering the information requested by a message box prompt.

#### Runtime Data
The MOD 2 Challenge's refactored code's runtime for 2018 stocks was discovered to be 28.9% slower than the Mod 2 lesson plan code; and, for 2017 stocks, the runtime was 4.6% slower than the Mod 2 lesson plan code.  It is my assumption that the position of the conditional code statements within the FOR loops to iterate over the rows as well as the dynamic input of the year into the macro via the message box makes a difference in runtime versus when the year is hardcoded within the macro code.


- Diane DeGette was the clear winner in this election contest winning 73.8% of the overall total vote with 272,892 votes.

![Module3PyPollTerminalResults](https://user-images.githubusercontent.com/79073778/111101194-c6292e00-851f-11eb-8dfb-4ebd3dfcc27a.png)


### Summary In Brief
This script can be used by the election commission across other multiple counties in the state. They would have to
append more data to the election data set.  They could also include more information by district, zipcode and party
affiliation to give a more robust view of the type of voters that support the candidates.


![Module3ChallengeVSCode](https://user-images.githubusercontent.com/79073778/111101610-a6463a00-8520-11eb-94a9-df091911fa8c.png)
