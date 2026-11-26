# Kinagin Farm Forms – Mobile App Prototype

A mobile-first Progressive Web App (PWA) for Kinagin Farm employees to record field data on their phones.

## Forms included
- Spray Records (Freshcare-style fields)
- Insect Reports (trap counts + photo)
- Irrigation Records
- Fertigation of Blocks
- Harvesting Records
- Diesel & Petrol

## Cloud Sync (Google Sheets)
1. Create a Google Sheet and open Extensions → Apps Script.
2. Paste the code from `GoogleAppsScript.js`.
3. Deploy as Web App (Anyone can access).
4. In the phone app, open **Cloud Sync Settings** and paste the Web App URL.
5. New records will appear in the Sheet (and stay saved on the phone).

## How to use on a phone
1. Open the GitHub Pages link.
2. Tap the browser menu → “Add to Home Screen” / “Install App”.
3. Use it like a normal app. Data is stored locally on the phone (works offline after first load).

This is a working prototype. It can be extended with more forms as needed.
