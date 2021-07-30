# Portfolio Tracker :chart_with_upwards_trend:

### Track your crypto portfolio using Google Sheets and CoinMarketCap.  

You provide coins tickers and quantities and the script will fetch prices and calculate totals and your portfolio allocation.  
You can define multiple currencies so you can see your portfolio value in USD, GBP, BTC etc..  

Similar to the following :point_down:  

![image](https://user-images.githubusercontent.com/1867877/127680438-4cc9cba5-a70e-49f9-8d0e-6c0968e272ec.png)


The advantages of doing this in a spreadsheet is that: 
  * :fire: you can very easily play with the numbers to run scenarios (what if bitcoin is $100K?)
  * :fire: it's easier than in Delta or Blockfolio/FTX to see everything at once
  * :fire: you can print it (why would you do that, though?)

:sparkles: As an added bonus, every time the prices are taken, a snapshot of your portfolio is added to a "history" sheet.  

# Setup (you only do this once)
0. Get a **free** API key from CoinMarketCap. This is where we get the price data from.  
Get it here: https://pro.coinmarketcap.com/signup

1. Create a new sheet in Google Sheets.  
  
![image](https://user-images.githubusercontent.com/1867877/127680160-3b18260c-ef68-4c54-bc23-7bae0baff0cf.png)

2. Click on Tools/Script Editor.  
  
![image](https://user-images.githubusercontent.com/1867877/127680212-2a0562a2-e4e6-4967-8001-98668e9bd994.png)

3. Remove the code in the Code.gs file and copy the code in https://github.com/bokobza/portfolio-tracker/blob/main/portfolio.gs  
After that, click "Save project".  
  
![image](https://user-images.githubusercontent.com/1867877/127680237-f9c4117d-5f1f-4999-94eb-093bcf32007b.png)
![image](https://user-images.githubusercontent.com/1867877/127680259-4b81b17f-fc50-48f2-9cc1-7a3f9a4a98d0.png)

4. Review the permissions process (this is all internal to you).
  
![image](https://user-images.githubusercontent.com/1867877/127680297-52e2b957-31e4-43a8-9741-cc4989d4b442.png)

5. Now head over to the "portfolio" sheet and add the coins and quantity you want to track.  
  
![image](https://user-images.githubusercontent.com/1867877/127680333-2f15e9f5-cf30-4ee9-a550-ffd686131c29.png)

6. Go to the "settings" sheet and copy the API key you got from CoinMarketCap.  
  
![image](https://user-images.githubusercontent.com/1867877/127680350-cd5f2fc7-eaa6-42c5-9b7a-dfe3ca7d05ad.png)

7. Set up is done. Now click on the "Refresh" button to compute all the data.  
  
![image](https://user-images.githubusercontent.com/1867877/127680402-5a3f4a56-0f31-4dcf-8647-fdc4a90ecf7d.png)

8. et voilà !  
  
![image](https://user-images.githubusercontent.com/1867877/127680438-4cc9cba5-a70e-49f9-8d0e-6c0968e272ec.png)
