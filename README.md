# amzGoogleApi
Python applications to evaluate Amazon drivers' efficiency utilizing Google API

- Application discription

About: This app uses the google sheets API to give drivers rating each day based on an algorithm and provide useful data for admin

Order:
 -  googleAPIDailyPerformance.py 
 -- Takes in google sheets id and date to give rating for each date

 - googleAPIFinal.py
 -- Takes in the google api and puts data in csv output: output.csv

 - pdftoData.py or simplifyAMZ_Local.py
 -- Takes in pdf weekly score and puts to csv output: nam_rating.csv

 - weeklyRating.py
 -- Takes in the csv MSCO and csv from AMZ to make one csv takes the average output: final_rating.csv

 - sort.py
 -- sorts the finalrating.csv output: sort_final_rating.csv
 
 - googleAPILoad.py
 -- loads the weekly data csv into the google sheets
