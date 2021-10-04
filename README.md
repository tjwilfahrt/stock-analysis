# VBA of Wall Street

## Overview of Project: Our friend, Steve, has asked for help analyzing a group of energy stocks to diversify his parents’ investments portfolio. Although they asked Steve to research a particular company, DAQO (stock symbol: DQ), he is certain VBA macros can assist analyzing all his stock suggestions simultaneously.

### Analysis: In addition to analyzing DQ’s performance, Steve has selected eleven other “green” stocks to include in his investment strategy. We analyzed two relatively large sets of data, “2017” and “2018.” To make for efficient analysis, our refactored macro prompts a user to specify a year to analyze (see string and user InputBox) and to activate an Excel sheet to populate a list of performance data in a legible format. The macro loops through Steve’s preferred stocks, their starting and ending prices, and displays information pertinent to his analysis. The macro also sorts data my Ticker variable, calculates daily trade volumes, and generates a return percentage to help Steve gauge which stocks he should include in his portfolio strategy.

![Screen Shot 2021-10-03 at 7 16 25 PM](https://user-images.githubusercontent.com/90878939/135778391-821fa3cb-2ec7-4013-8cd2-98dcdc22102b.png)

![Screen Shot 2021-10-03 at 7 16 10 PM](https://user-images.githubusercontent.com/90878939/135778395-a151990c-8475-4d9a-b95b-89421ea1fc8f.png)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/90878939/135778136-9eb98557-4649-48cc-bf6e-ca9535ab3eef.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/90878939/135778139-6d8e2144-4e4e-4f87-89f3-84027c99583e.png)

![Screen Shot 2021-10-03 at 7 18 18 PM](https://user-images.githubusercontent.com/90878939/135778421-c18e21d0-0051-4dcc-9023-ea252ada1056.png)

![Screen Shot 2021-10-03 at 7 18 50 PM](https://user-images.githubusercontent.com/90878939/135778433-80124b32-4cd2-4cf0-a684-cfc104a15217.png)

![Screen Shot 2021-10-03 at 7 19 38 PM](https://user-images.githubusercontent.com/90878939/135778483-72bbf62d-75e1-4800-9f93-75b78cae2e30.png)

![Screen Shot 2021-10-03 at 7 20 03 PM](https://user-images.githubusercontent.com/90878939/135778485-05b46c1a-ab17-45b8-8ddd-a8b7981e7b93.png)

### Challenges: Refactoring Steve’s macro was a significant challenge. I integrated “And” statements to integrate efficiency, which only resulted in erroneous results; VBA generated accurate data more quickly without these ‘And’ strings. I leaned heavily on my TAs, LAs, office hours, and study groups to carefully review my strings from top to bottom. 

## Results: Steve’s stocks performed much better in 2017. His parents were right: DQ performed the best amongst all companies. Execution times for analyzing 2017 and 2018 data was similar and .1 seconds to process. Using refactored code cut analysis time significantly simply because all subroutines are run all at once, rather than individually. It is also helpful to input a year on the front end of the analysis to prevent having to alter strings throughout each macro.   

## Summary: Refactoring code has advantages and disadvantages. Advantages include a more concise, easier to follow, and shorter routine that generates a result quicker. However, refactoring code takes time to process, which may defeat the very purpose the activity. As I found, refactoring Steve’s code led me down a path that I could not solve without the help of my peers. Our original analysis required multiple steps with our first macro. Although each macro processed quickly, users could identify what VBA was processing in a step-by-step fashion. We’d have to scroll down to through each macro to get the full picture. Our refactored macro also asked which data set to analyze automatically, whereas our original code defaulted to 2018 data. Refactor code also formatted data uniformly.
