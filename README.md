# Green Lady cider - barrel and bottle tags

Records live in a Google Sheet (tabs Barrels, Blends, Log, Ratings) served by the Apps Script web app in `apps-script/Code.gs`. The page reads and writes through it.

- `#B07` - barrel page (NFC tag + QR sticker on the barrel). Add a note; More: Reading, Yeast, Bottled, Emptied (PIN number).
- `#26A-B07-031` - bottle page: the story of the bottle + 👍 / 👌 / 👎 rating (no PIN).
- `#blend/26A`, `#league`, `#labels/26A-B07/23` (bottle QR labels, A4 21-up), `#labels/barrels`.

Tag writer: `python tools/write_tag.py B07` (never locks tags).
