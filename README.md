# HLK OBS Database

HLK OBS issue tracker built with Streamlit. This package contains source code
and dependency information only; it does not include the local database,
Outlook messages, or uploaded attachments.

## Run on Windows

1. Install Python 3.10 or later.
2. Open a terminal in this folder and install dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

3. Start the app:

   ```powershell
   python -m streamlit run HLK_OBS_Database.py
   ```

The app runs locally and uses classic Outlook installed on the same computer.
Uploading this source to GitHub stores the code; GitHub does not run this
Streamlit application.

## Data and privacy

The SQLite database and uploaded files are created locally when the app runs.
They are deliberately excluded from this source package. Review changes and
confirm your organization's sharing policy before uploading source code to
GitHub.
