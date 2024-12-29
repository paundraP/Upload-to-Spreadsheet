# Uploading Excel File to Spreadsheet use Golang

I created this project to make some of my work easier. 

## Prerequisite 🧰

First, you need to set up your Google Sheet API access. If you do not know how to do this, make sure you follow this step:

1. Go to the [Google Console] (https://console.cloud.google.com)

2. Create a new project or use your existing project

3. Go to APIs & Services -> Enable APIs & Services -> Search for and enable the Google Sheets API.

4. Go to Credentials and Enable it for Service Account, follow the step to create it.

5. Then download the credentials.json. However, in some cases the credentials.json is not downloaded automatically. To make sure you have the credentials.json, follow these steps
    
    - Go to IAM & Admin
    - Select Service Account
    - Select the account you've created
    - Go to Keys and create a new key
    - Select json and create. 

6. Save the credentials.json and the email of the service account

Then go to google sheet and create a new one. Click the Share button and add your service account email to the spreadsheet. And copy the link from your spreadsheet.

## How to use🚀

1. Clone this repository
```bash
git clone xxx
```
2. Go inside the folder
```bash
cd xxx
```
3. Install the package dependency for this project
```bash 
go mod tidy
```
4. Copy the .env.example and add your spreadsheet id from the link you've copied before.

if your link is like this: 
https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwxyz/edit?usp=sharing

the spreadsheet id is the part between /d/ and /edit. So base of the example above, the spreadsheet id is **abcdefghijklmnopqrstuvwxyz**

5. After that, place your credentials.json in this folder too

6. You can try to upload your excel file to spreadsheet by running
```bash
go run main.go --uploadfile /path/to/your/excel.xlsx
```