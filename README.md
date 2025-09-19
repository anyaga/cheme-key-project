# General Description

This is a project that is hosted on Google Appscript and is compatabile with severle google products. This project was completed to asset a Joint IT Department in Carnegie Mellon University's College of Engineering

The purpose of this project is to determine who should have key access. Keys are given to students, staff, faculty, and adminstration. Once keys are given, we collect the individual's name, the individual's ids, the key number, the room number, the given date, and the expiration date. Key information is collected on Google Spreadsheets and a Google Form.

The goal of this project is to parse key information to determine keys that are expired and keys that will expire soon. Keys that will expire soon are displayed on a main spreadsheet. With the press of a button, the faculty adminstrator can contact these individuals. The same is true for expired keys. If too many keys to the same room are expired and not collected, a notification will appear to suggest re-keying the room to ensure the security of the space.


# Authorized User Sheet

This page is used to determine who has access to the 'Key Tracking Sheet'. This is done through automation, direct editing of the spreadsheet and 
the press of a button (connected to a function)

# Key Tracking Sheet

This folder process inputs, analyzes them, and then produces new organized data as an output



# clasp commands (google apps script)
clasp login

clasp pull
clasp push

clasp open

clasp run <functionName>


# Git commands

git add .
git commit -m "---"
git push
git checkout branch_name


