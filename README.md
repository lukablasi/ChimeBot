# ChimeBot
Amazon Chime Webhook that can update a chat with a shift performance update. Running on NodeJS. The App once started pulls Excel file from sharefolder with a shift plan for a day to exctract data from it and save it in it's local directory. Based on that data shift performance is calculated. The app will try to pull the file every day at 08:50 and 09:50. for a second time in case preparing the shift plan was delayed. <br>
At 09:00 app will pull data from PPR with current shift performance metrics and will calculate if hourly plan is being achieved. That will continue to happen every 30 minutes until the end of the shift. <br>
At the end of the shift the app will delete saved shift plan from local directory. It would overwrite an existing file if there was one while ChimeBot is trying to save new shift file with a new data, but this step is necessary to avoid a sitution when in a morning creation of a shift plan is delayed and the bot is posting data from a previous day until new one is created.

![](https://github.com/lukablasi/ChimeBot/blob/main/screenshots/flowbot.PNG)

## Technologies Used
<ul>
  <li>JavaScript</li>
  <li>NodeJS</li>
  <li>Express</li>
  <li>Cron</li>
  <li>Playwright</li>
  <li>XLSX</li>
</ul>

## 
