# stock-analysis
Analysis of Green Energy Stock

# Kickstarting with Excel
    When starting this project the main objective was utilizing Microsoft Excel to create an analysis and add some VBA code to automate the steps that may repeat diring the proces.

## Overview of Project
    This project is an analysis of historic data of 12 green stocks to determine how appropriate it may be for the clients "Steve's Parents" investment plans. We are processing historic data and provide athe Total Daily Volume and rate of return for each of the stocks in our dataset.

    ![initial analysis results](https://raw.githubusercontent.com/jhonpire/stock-analysis/main/Resources/initial_analysis_results_2017.png)(https://raw.githubusercontent.com/jhonpire/stock-analysis/main/Resources/initial_analysis_results_2018.png)

### Purpose
    The purpose of this Project is to make facilitate the processing and analysis of data for Steve's parents investments plans by automating the process of selecting, filtering, calculating and formating of information to turn thousands of lines of data into easy-to-read information, and that way make their selection easier.

## Analysis and Challenges
    Analyzing the dataset received from Steve presented with some challenges, the firs one being that the data was  organized by Stock name and not by date making the process to be longer as the code had to run many times over the complete dataset to generate the information of each stock, making the code to run for longer and consume more computer resources. Another challenge is that there was only 2 years of historic data, meaning that the analysis although could compare the performance from one year to the next, would have been much better to have information to work with and find more patterns.

### Analysis of Outcomes Based on Launch Date
    The outcomes for the behavior of the stocks on the market meant that the year 2017 was much more profitable across the Green Energy industry. Since our analysis is focused on the stocks for Daquo or "DQ", we can conclude that the fall on the value of the stocks means that it wouldn't be a good idea to invest on that specific stock.

### Analysis of Outcomes Based on Goals
    The analysis performed of the data provided showed that investing on Green energy is not a good idea for Steve's parents as we discovered that 10 out 12 of the stocks included in the sheet, presented negative numbers for the year of 2018. Compared to the previous year 2017, wher only one stock had been on negative with 11 showing returns of investement.

### Challenges and Difficulties Encountered
    With the code writen to perform the analysis, we noticed that the way it had been factored made it take longer to perform the analysis and generate results. For that reason we decided to refactor the code and change the order in which the information would be processed to save time and resources. Here we show a screen capture presenting shorter times to run compared to the previous version.
    ![Refactored analysis results](https://github.com/jhonpire/stock-analysis/blob/main/Resources/refactored_analysis_results_2017.png?raw=true)(https://github.com/jhonpire/stock-analysis/blob/main/Resources/refactored_analysis_results_2018.png?raw=true)

## Results

What are two conclusions you can draw about the Outcomes based on Launch Date?
    - Based on the Launch date, it seems like Green Energy stocks have been losing value drastically compared to previous years, not only on the targeted stock "DQ" But the majority of the stocks available in our dataset.

What can you conclude about the Outcomes based on Goals?
    - Based on the goal to determine if DQ was a good stock to invest, we have determined that it was not a viable stock to place any investments into.

What are some limitations of this dataset?
    - This dataset only contains the information of 2 years of performance records which in most cases may be too little information to make this type of decision.

What are some other possible tables and/or graphs that we could create?
    - According to the type of information available, we could create ***Bar Graphs*** to compare side by side the performance of each stock based on the year of investments on the stock market.