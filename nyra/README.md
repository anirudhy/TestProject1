# Nyra's Birthday Invite 🎈

A one-page birthday invitation website. Plain HTML, CSS and JavaScript, no build step.

## Edit the party details

Open `index.html` and change the `PARTY` block near the bottom:

```js
const PARTY = {
  name: "Nyra",
  start: "2026-11-14T15:00:00-08:00",
  end:   "2026-11-14T18:00:00-08:00",
  placeName: "Our Home",
  address: "Seattle, WA",
  notes: [ ... ],
  rsvp: "Please RSVP by text to Anirudh"
};
```

The date card, countdown, "Add to Google Calendar" and "Get directions" buttons all update from it.

## Run it locally

```bash
cd nyra && python3 -m http.server 8000
```

Then open http://localhost:8000 in your browser.

## Live site

https://anirudhy.github.io/TestProject1/nyra/ (served by GitHub Pages from the master branch)
