## Stockananlysis using VBA##

**Overview of Project:**

We are trying to do stock performance analysis for some leading green energy companies for the year 2017 and 2018, We have written separate codes to run performance for each year, However the challenge is to refactor the code to increase flexibility and efficiency.

**Results:**

Refactoring the code helps to extend this analysis, add data for 2019 ,2020 or any years in the future and continue the analysis with the same code with minimal modifications, it helps increase the efficiency because all the required calculations and loops are in one code instead of multiple codes. refactoring improves objective attributes of code that correlate with ease of maintenance, if I Have to rewrite a certain part of the code or add functionality it’s easy to do so.
When refactoring I understood the code better, adding comment along helped me with final code check and corrections as well, refactoring helped me understand design decisions, for example when I declared the following variables “Tickerindex,Tickervolume,Tickerstarprice and Tickerendprice “ I was forced to visualize the flow of code and how each loop will use these variables.

![code](images.zip/images/loop1.PNG)

![code](images.zip/images/variable2loop2.PNG)

![code](images.zip/images/ifinloop.PNG)

![code](images.zip/images/loop1.PNG)

Refactoring helped me understand the loops and nested conditions better as well, The first "for" loop named “i” makes space for volume ,start and end price for 12 stocks and initialize their value as 0 , That way when we add values we know we started from 0 ,no room for miscalculations. The second "for" loop name “inputrow” loops through the row until end of the data, and add the required values for all 12 stocks.

**Summary:**

My refactored code for 2017 and 2018 runs in 1.71 seconds and 1.67 seconds respectively, we have analyzed 12 stocks from 2 years under 4 seconds . below are my runtimes 

![code](images.zip/images.zip/images/runtime2017.PNG)

![code](images.zip/images.zip/images/runtime2018.PNG)

**Here are some insights from the anlysis**

    1.DQ, SEDG ,ENPH FSLR did very well in the year 2017
    2.RUN, ENPH are the only 2 stocks did well in 2018

![code](images.zip/images/2017stocks.PNG) 

![code](images.zip/images/2018stocks.PNG)

I started this analysis to advice whether to invest in DQ green energy company, with this code we could see that DQ stock returns for 2018 Is -62.6% and it was 199.4% , this looks like a very volatile stock . I would prefer and advice investing in a much more stable performing stocks as we are making this suggestion for a retired couple.

ENPH stock performance is more stable the volume has increased in 2018 still the returns are positive for both 2017 and 2018. Calculating the gain or loss on an investment as a percentage is important because it shows how much was earned as compared to the amount needed to achieve the gain, a positive returns percentage is what we are looking for and ENPH is the stock to invest for stable growth and returns.


