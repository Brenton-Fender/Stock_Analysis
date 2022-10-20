# Stock_Analysis

#Green Stocks Analysis

##Overview
The purpose of this project is to help analyze some stock data for my friend Steve. Steve would like to help his parents diversify their investments in green energy stocks. Steve has provided me with an Excel worksheet that has a list of green stocks accompanied by two years of trading data. With the data Steve provided, I will write a VBA script that allows Steve to quickly analyze the volume and yearly returns with the click of a button.

##Results
By using a VBA script with embedded formulas, Steve can now view how much a stock is trade along with the yearly returns for 2017 and 2018. The script saves Steve from doing manual calculations and guarantees the accuracy of the calculations. The way in which I wrote the script will allow Steve to add more data in the form of additional years and analyze it at the push of a button.  With data provided we can see that only ENPH and RUN managed a positive return for both 2017 and 2018. Only TERP managed a negative return for both years. By looking at the pictures I have included, you can see that 2017 was an overall good year for this list of stocks and 2018 was an overly negative year. This makes it notable when a stock performed well or poorly through both periods.


######2017 Data Refactored Code
![VBA_Challenge_2017](https://user-images.githubusercontent.com/105960365/196993813-282cae86-6b71-4d4f-97fd-db360d6c4da1.png)


######2018 Data Refactored Code
![VBA_Challenge_2018](https://user-images.githubusercontent.com/105960365/196993468-1692a853-7c7e-4556-89e9-6130af91b255.png)


######2017 Data Original Code
![Original_2017](https://user-images.githubusercontent.com/105960365/196996928-3c7e24c2-0e33-4042-b986-e1ffa97eef66.png)


######2018 Data Original Code
![Original_2018](https://user-images.githubusercontent.com/105960365/196997017-a2e6af2e-98cd-4f8f-88fc-89a1e911ce47.png)


##Code Performance
The performance of the refactored code was significantly better than the original code. The refactored code performed the same analysis in approximately 13% of the time as the original code. In a matter of actual time saved, three tenths of a second is not considerable. However, if this project was scaled up to include more stocks and more years of data, the time savings could begin to become significant.

##Summary
The major advantage of refactoring this code was that I already had many of functions and formulas completed and I could simply copy them to my new script. This provides a clear time savings for completing the project. A disadvantage to refactoring code is that if you have stepped away from a project for any amount of time, the reimmersion can be confusing. This confusion can be overcome by coding with detailed notes of what the code is accomplishing, using clearly defined variables, and avoiding the use of "magic numbers." By refactoring this code, the performance was sharply elevated. The new performance of the code will show Steve that my work is scalable to larger projects with more data. The only disadvantage of refactoring this code is that the instructions for the original code did not calculate the yearly returns correctly. In order to have matching data for this project, I had to correct the original code and take new screenshots.
