# stocks-analysis
using VBA to help Steve in stock analysis
# VBA Of Wall Street

## Overview of Project

The purpose of this is to help new finance graduate, Steve in the analysis of various green energy stock data. His parents are looking to invest in successful stocks and have become his clients in order for him to help determine the best stocks to invest in based on the performance of these stocks over certain time periods (particularly the years 2017 and 2018)

## Results

VBA_Challenge_2017.png https://github.com/SNwokolo/stocks-analysis/blob/ff06299af32d287d0a83e613aa2da80a1c35ea71/VBA_Challenge_2017.png
https://github.com/SNwokolo/stocks-analysis/blob/ff06299af32d287d0a83e613aa2da80a1c35ea71/VBA_Challenge_2018.png
The above images show one of the results of running the code for the "All Stocks Analysis" worksheet in the VBA_challenge.xlsm file. For the year 2017, the code ran for 0.1953125 seconds and for 2018, it ran for 0.2265625 seconds.

From the results of the analysis run, 2017 was a more successful year for majority of the green energy stocks in general. The only stock with a negative return that year was TERP (-7.2%). In 2018, majority of the stocks had poor returns with the exception of ENPH and RUN which had returns of 81.9% and 84.0% respectively. Given that these two stocks (ENPH and RUN) also thrived in 2017 with returns of 129.5% and 5.5%, from the looks of the data these are two stock options that Tom should have his parents look into investing if they wish to gain profits.

## Summary

### Advantages and Disadvantages of refactoring code
**Advantages**
Some advantages of refactoring code include: 
- it makes code more clean, manageable, readable and easier to reuse[^1][^3].
- it could help the program/script run faster[^2].

**Disadvantages**
- it can be time consuming[^2][^1].
- it could potentially lead to more bugs and errors in the script[^2].

**Relating the pros and cons to the original VBA script:**
With regards to this particular script, one obvious pro with refactoring the code was the fact that the scripts ran for a shorter period of time in the refactored script when compared with the original scripts for both 2017 and 2018. The refactored script in this analysis was also more efficient in the sense that the resulting data was more manageable and clean.
For cons, due to the lengthier nature of this new script, it was harder to detect errors when they occured which in turn made the process of refactoring more time consuming as it was difficult trying to locate some of these errors despite how minor they seemed to  be. One example of this was noticed when one of the loops created wasn't properly ended with a 'next'. it was a little difficult finding which particular loop the problem was in due to how many loops were present in the script.

## References
1. Ribeiro, Edward. “What Are the Pros and Cons of Refactoring?” Quora, 2013, https://qr.ae/pvOtb2. 
2. Ershad, Gul Md. “Pros and Cons of Code Refactoring.” C# Corner, 9 Jan. 2017, https://www.c-sharpcorner.com/article/pros-and-cons-of-code-refactoring/. 
3. Pandey, Be Wake. “What Is Code Refactoring and What Is It Used for?” Quora, 2018, https://qr.ae/pvOtWY. 
