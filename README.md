# Double Booking Checker Excel Add-in

This Excel add-in helps you check for double bookings in your schedule by analyzing date and time columns in your spreadsheet.

## Setup Instructions

1. Create a new GitHub repository named 'double-booking-checker'

2. Push your code to the repository:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/[YourUsername]/double-booking-checker.git
   git push -u origin main
   ```

3. Enable GitHub Pages:
   - Go to your repository settings
   - Navigate to 'Pages' section
   - Under 'Source', select 'main' branch
   - Click 'Save'

4. Update manifest.xml:
   - Replace all instances of `[YourGitHubUsername]` with your actual GitHub username

5. Build and deploy:
   ```bash
   npm install
   npm run build
   git add .
   git commit -m "Build for production"
   git push
   ```

6. Test the add-in:
   - Download the manifest.xml file
   - In Excel, go to Insert > Office Add-ins > Upload My Add-in
   - Select the manifest.xml file
   - The add-in should now appear in Excel

## Development

To run the add-in locally:

1. Install dependencies:
   ```bash
   npm install
   ```

2. Start the development server:
   ```bash
   npm start
   ```

3. In Excel, upload the manifest.xml file as described above

## Features

- Automatically detects date/time columns in your spreadsheet
- Checks for overlapping time slots
- Provides detailed conflict information
- Modern UI with Fluent UI React components