# Challenge
## jmerkel_Module_02
UC Berkeley Bootcamp - Module 02 - VBA - Joe Merkel

After creating the title and header rows, the program declares the 4 arrays with a size of 10. Since it is unknown how many companies will be analyzed, the arrays are declared via ReDim in order to be dynamic. Additionally, a ticker index is declared and initialized to 0. A majority of the logic is performed within 4 IF statements in the row iteration loop.

The first if statement checks to see if the tickerIndex is outgrowing the array. If the ticker is bigger than the array size, then the arrays are doubled in size while preserving the contents.

The next IF statement checks to see if the ticker is new. If it is, then the ticker name & start prices are recorded in their respective arrays. Otherwise it moves to the next line.

The 3rd if statement verifies that the current row (ticker name) matches the current indexed name. If it does, the volume is added to the running total in the volume array.

Finally, the last if Statement checks to see if the current row is the last line in the sorted list. If so, it records the end price in the End Price array and outputs the indexed values from the 4 arrays into the output columns.

The final section of the code formats the output and displays a Message Box that shows that the macro is complete.
