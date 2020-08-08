# VBA Challenge

## Overview of project
We need to refactor the macro just in case Steve wants to have more than just 12 stocks. Refactoring the code is making it more efficient and potentially faster. This can include skipping certain steps from before, and combining other steps as well. I refactored the code to where it runs faster and more efficiently. 

##Results
For the code that is not refactored, the timer says it takes around .8 seconds for the macro to run for years 2017 and 2018. The refactored code is around four times faster, taking only .2 seconds to run the refactored macro. Below are two pictures. The first one is the output and the timer that says how long it took to run the regular code. The second picture shows how long it took to run the refactored code. 
### Timed Results
![2018 no refactor vba chall](https://github.com/JoelS-Pebbles/stock-analysis/blob/master/2018%20no%20refactor%20vba%20chall.PNG)
![2018 refactor vba chall](https://github.com/JoelS-Pebbles/stock-analysis/blob/master/2018%20refactor%20vba%20chall.PNG)

### Code example
Below is the code for the macro that was not refactored. 

If Cells(j, 1).Value = ticker Then
        totalvolume = totalvolume + Cells(j, 8).Value
    End If
    
Below is the refactored code that does the same thing as the example above. 

tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(j, 8).Value

Both of the code examples are finding the volume for the current ticker. I belive the refactored code is faster at proccessing the macro because it has less arguments and variables in the code, like the example provided above. 


## Summary
The refactored code was much faster in proccessing the macro. This is especially useful if Steve wanted to add some more stocks to analyze. A disadvantage to refacoring code is you could make a mistake and forget how you got the code working in the first place. Or you might present new bugs that could mess up the code. It was easy to tell how much faster the code ran thanks to the timer that we put in both macros. But with refactoring code, I moved some things around and it too several tries before I got the lines of code where they were supposed to be. 
