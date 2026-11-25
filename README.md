# URL Checker Pro — By ENG Ahmed Bahgat

A desktop URL & redirect-chain checker built with Python, CustomTkinter, `requests`, `BeautifulSoup4`, and `Playwright` using installed system browsers (**Microsoft Edge** or **Google Chrome**).

## Features
- **100% English UI:** Clean English interface with zero RTL text reversal issues.
- **Thread-Safe System Browser Engine (`ThreadLocalBrowserChecker`):** Avoids `Cannot switch to a different thread` by isolating Playwright instances per worker thread and supporting `msedge` / `chrome`.
- **Smart Mode (`Smart HTTP + System Browser`):** Automatically escalates to the system browser on `403` / `Access Denied` / bot-gate pages or JS/Meta redirects.
- **Multi-Format Export:** Exports reports to `.xlsx` (`All Results` & `Errors Only`), `.csv`, and `.txt` (Tab-delimited).
- **Automated GitHub Releases:** Builds `URL_Checker_Pro.exe` via GitHub Actions and publishes releases on pushes to `main` or `v*` tags.
