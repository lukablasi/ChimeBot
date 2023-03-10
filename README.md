# ChimeBot
Amazon Chime Webhook that sends messages out to a chat room with a shift performance update. Running on NodeJS. The App once started pulls Excel file from sharefolder with a shift plan for a day to exctract data from it and save it in it's local directory. Based on that data shift performance is calculated. The app will try to pull the file every day at 08:50 and 09:50. for a second time in case preparing the shift plan was delayed. <br><br>
At 09:00 app will pull data from PPR with current shift performance metrics and will calculate if hourly plan is being achieved. That will continue to happen every 30 minutes until the end of the shift. <br><br>
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

## Set Up Instructions
<ol>
  <li>Download project files to your local directory.</li>
  <li>Make sure you have NodeJS installed on your machine</li>
  <li>Open a terminal in a folder where project's files were downlaoded and run command "npm install"</li>
  <li>If installation ends successfully run "node index.js"</li>
  <li>After those steps in a terminal you should see a message "Server is successfully running on port 3000".</li>
  <li>Leave the terminal open, it will print out a message in a console every time ChimeBot sends a chime out.</li>
  
  ![](https://github.com/lukablasi/ChimeBot/blob/main/screenshots/terminal.PNG)
  
</ol>

## How to use
The app will open a browser page and select required options from PPR. This process will not be seen by a user however can be if Playwriht's configurations are changed in "index.js". The ChimeBot will select desirable settings but those can be adjusted in a code.

![](https://github.com/lukablasi/ChimeBot/blob/main/screenshots/ppr.PNG)

Make sure your app has correct link to the place where shift plan is located. Make sure it is searching for a file with a correct name - by default it a current day in a format DD/MM/YYYY with a string "Shift Plan.xlsm". The app will not be able to find a file if file's name doesn't have an expected value.

![](https://github.com/lukablasi/ChimeBot/blob/main/screenshots/shiftplan.PNG)

Adjust your "planModifier" - it's a variable that calculates % of a plan that should be processed at a certain point in time. Example - on a snip below at 09:00 11% of the planned volume should be processed and at 13:30 that would be 56%. You can change a value of the modifier and adjust it to your shift trends.

![](https://github.com/lukablasi/ChimeBot/blob/main/screenshots/planmodifier.PNG)

Cron functions at the bottom set a time schedule for each function. Be default "main" function which is responsible for calculations and sneding messages out runs at every hour from 09 to 18 at each 00 and 30 minute. If you want to adjust it see Cron documentation for a detailed syntax.

![](https://github.com/lukablasi/ChimeBot/blob/main/screenshots/cron.PNG)

<h2>Created By</h2>
Lukasz Milcz <br>
