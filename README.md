# Refactoring VBA Code for Stock Analysis

### Overview of Project
Refactoring VBA code designed for analysis of a handful of green energy stocks so that it will work efficiently when performing the same analysis on the entire stock market.

#### Results

##### Initial VBA Code
The initial VBA code (Module 2) examined 12 green energy stocks, looking at their total daily volume and return on the chosen year. Here's a look at the chunk of code that performed this analysis:

![Initial Code](https://user-images.githubusercontent.com/107162310/174845059-afdb8db0-ff0c-4949-a31f-5dd76fccb472.png)

The worksheet included a button that allowed the user to type in the year they wanted to analyze (2017 or 2018), using the following code:

> yearValue = InputBox("What year would you like to run the analysis on?")

Also included in the code was some formatting, a button to run the analysis, and a button to clear the sheet to run it again. Finally, a timer was built to track the amount of time the code took to run:

> startTime = Timer
> endTime = Timer
> MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

##### Initial VBA Code Run Time
Using this timer, the initial code took 0.890625 seconds to process the 12 stocks for the year 2017 & 0.9257813 secpmds to process for the year 2018.

##### Refactoring for Improved Performance
The challenge was to refactor this code so that it ran faster and allowed for upscaling to include all the tickers on the stock market rather than just a dozen. 

##### Converting Unix Timestamps
The dataset also included the launch dates and deadlines (Launched_at & Deadline columns) for every project, however the data was provided in Unix timestamps. To make these usable for potentially valuable analysis, new columns were created (Date Created Conversion & Date Ended Conversion) and the Unix timestamps were converted to more traditional calendar dates using the following formula:

> =((([Unix_timestamp_cell]/60)/60)/24)+DATE(1970,1,1)

##### When Do Successful Theater Kickstarters Launch?
To suggest a launch date for a theater Kickstarter campaign, first I added a "Years" column to the dataset and filled it with data by using the YEAR() function to extract the year portion from the Created Conversion column. Then a pivot table was created, with filters set to the "Parent Category" & "Years" columns. After filtering to the "theater" category, the following line table was created:

![Outcomes_Launchdate_chart](https://user-images.githubusercontent.com/107162310/174155204-32587feb-81d9-4679-8f0b-792a60c6c7e0.png)

From this table, the following line chart was created:

![Theater_Outcomes_vs_Launch](https://user-images.githubusercontent.com/107162310/173885633-beb3bcf6-b89a-4dcf-9d55-eacc04a1d481.png)

From this data visualization, it's clear that the successful theater Kickstarter campaigns separate themselves from the failed or canceled campaigns starting in the spring months, and the difference in success becomes less stark as the summer stretches on into autumn. **Ideally, a theater Kickstarter should target the month of May for its launch date.**

##### What Goal Amounts are Most Successful?
To determine a goal amount that was more likely to be successful, I looked at twelve different goal ranges in a table on a new sheet. I used the COUNTIFS() function to pull the number of failed, canceled, and successful Kickstarters with a "plays" subcategory. Here's an example that I used to find the number of successful "plays" campaigns with a goal of less than $1000:

> =COUNTIFS(Kickstarter!$F:$F,"successful",Kickstarter!$D:$D,"< 1000",Kickstarter!$R:$R,"plays")

After that data was pulled into the table, I then used the SUM() function to total the number of projects for each of the twelve goal ranges. Finally, I calculated the percentages of the three different outcome possibilities for each of the goal ranges to produce this table:

![Goals Table](https://user-images.githubusercontent.com/107162310/173905038-5831b12a-065c-47d8-be77-4888a244b42b.png)

Pulling this table into a line chart helped to illustrate the goal amount that heavily leaned toward successful Kickstarter campaigns.

![Outcomes_vs_Goals](https://user-images.githubusercontent.com/107162310/173905476-d743c224-43dc-4617-8c1c-df66e1ff902e.png)

The Percentage of Successful "plays" campaigns is highest on the lower range of goal amounts, and the delta between the successful and failed campaigns also happens to be widest at this same point. **Ideally, a "plays" Kickstarter should target a fundraising goal of less than $5,000.**

##### Challenges with the Analysis
I found it challenging to fill the Goals Table using the COUNTIFS() function initially. I've used plenty of COUNTIF() functions with several criteria ranges, but this was the first time I had a shifting goal amount to keep in mind. Eventually I realized that I had to include an extra criteria to get it to work properly for any goal range that had two end points. For instance:

> =COUNTIFS(Kickstarter!$F:$F,"successful",Kickstarter!$D:$D,">999",Kickstarter!$D:$D,"<5000",Kickstarter!$R:$R,"plays")

Unlike the earlier code snippet, this goal range has both a greater than figure to track for and a lesser than figure.

In checking my work, I also got lost in the number of filters that went into analyzing the dataset and for a while there I could not get the full dataset to appear and was worried I had done something to remove chunks of the data. Eventually I remembered the Clear Filters button.

#### Results

##### Launch Date Conclusions
The analysis points to two conclusions based on launch date:

1. Most successful theater Kickstarters are launched in spring (May) or early summer
2. Fundraising campaigns launched at the end of the year have almost failed as many times as they've succeeded

##### Goals Conclusions
The fundraising goal of the campaign should be less than $5,000. Not only were the highest percentage of successsful campaign goals at this level, but the lowest percentage of failed campaigns also were set below $5,000.

##### Dataset Limitations
While the dataset allows for certain conclusions to be drawn, it is missing potentially critical information that could help explain the success, or lack thereof, of fundraising campaigns. For instance, it includes the overall amount pledged and the total backers, which allowed for the calculation of an average pledge, but it's missing data on the backer rewards. Is it possible that campaigns that offered perceived lackluster rewards (either from monetary value or something else) failed at higher rates? Survey data on backers' or potential backers' opinions about the backer rewards could help us draw conclusions here.

There's also no information on what sort of marketing the various campaigns used to attract backers. We know that campaigns in May were more successful, but why? Did they make their marketing pushes in that month? On the whole, were successful campaigns more likely to be savvier marketers? There's no way to know here, but survey data on what attracted backers to a campaign, or pushed them away from one, would be helpful.

##### Possible Further Analysis
For further analysis, a table could be made looking at the average pledges among theater campaigns with various outcomes. If a certain amount is more prevalent among successful campaigns, any future campaigns could target that amount as their entry level backer reward.

We've analyzed the launch month of projects but with our dataset we could also look to see if the length of time a campaign runs has any potential bearing on its results. Perhaps campaigns with shorter lengths create a sense of urgency in potential backers, leading to better success rates. Analyzing the data through a table and/or chart could tell us if that were the case.
