# Food Court Line Waste Log

A tablet app for chefs' end-of-night waste count at Pechanga Fried Chicken, Pronto, Little Wok and Agave. After every outlet is counted, it emails one combined report to management and logs the data to a Google Sheet.

![Sample email](docs/sample-email.png)

- **App:** `index.html` is a single file with no build step, hosted on GitHub Pages.
- **Email + logging:** `apps-script/Code.gs` is a Google Apps Script web app.
- **Setup:** see [docs/SETUP.md](docs/SETUP.md).
