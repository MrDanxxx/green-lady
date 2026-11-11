# Green Lady cider - barrel tags

Each barrel has an NFC tag holding one link: `https://mrdanxxx.github.io/green-lady/#B1`.
Scanning it opens that barrel's current fill: apples and where they grew, starting gravity and pH, yeast, juicing date, readings.

Edit `data.js` to update. New vintage each autumn: add a batch (e.g. `GL27-1`), move the old batch ID into the barrel's `history`, point `batch` at the new one. The tag never needs rewriting.
