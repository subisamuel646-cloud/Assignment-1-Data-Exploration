# Assignment-1-Data-Exploration
Excel Formulas

1) total price of products : SUM : =SUM(range)
   No. of products : COUNT : =COUNT(range)
   Average price of products : AVERAGE =AVERAGE(range)

2) Determine the minimum price among all products : MIN :=MIN(range)
   Find the maximum price among all products : MAX : = MAX(range)

3) Using an IF function, create a new column named Price Range to categorize products with a price greater than or equal to $500 as 'High Price' and others as 'Standard Price'
   Column inserted "Price Range"
   IF Function : =IF(condition, value_if_true, value_if_false)

4) Calculate the total price for products in the 'Electronics' category using the SUMIF function.
   SUMIF : =SUMIF(range, criteria, [sum_range])

   Determine the count of products with a price less than $100 using the COUNTIF function
   COUNTIF : =COUNTIF(range, criteria)

5) Text Formatting - LEFT, RIGHT, MID:
   Create a new column named Day with the first 2 characters of each 'Product ID' using the LEFT function.
     -----> Added column "Day" : LEFT() : =LEFT(text, [num_chars])

   Create a new column named Country Code by extracting the last 2 characters from the 'Product ID' column using the RIGHT function
     -------> Added column "Country Code"  : RIGHT() : =RIGHT(text, [num_chars])

   Create a new column named Month by extracting 4th to 6th characters from the 'Product ID' column using the MID function
     -------> Added column "Month" : MID() : =MID(text, start_num, num_chars)
