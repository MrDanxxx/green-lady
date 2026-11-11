// Green Lady cider - barrel records.
//
// Each NFC tag holds ONE link ending in the barrel's ID, e.g.
//   https://mrdanxxx.github.io/green-lady/#B1
// The tag stays on the barrel for life. Each autumn, add a new batch below
// and point the barrel at it - no need to rewrite the tag. Earlier batches
// stay listed under "Previous fills".
//
// Dates are YYYY-MM-DD. Gravity like 1.052. Anything you don't know yet,
// just leave out - the page only shows what is filled in.
// NOTE: this page is public. Describe tree locations loosely
// ("Home orchard", "Neighbour's tree"), not full addresses.

const SAMPLE_DATA = true;   // set to false once the real details are in

const BARRELS = {
  B1: { name: "Barrel 1", size: "25 L", batch: "GL26-1", history: [] },
  B2: { name: "Barrel 2", size: "25 L", batch: "GL26-2", history: [] },
  B3: { name: "Barrel 3", size: "30 L", batch: "GL26-3", history: [] },
};

const BATCHES = {
  "GL26-1": {
    vintage: 26,
    juiced: "2026-10-04",
    apples: [
      { variety: "Dabinett",        location: "Home orchard",       share: "40%" },
      { variety: "Bramley",         location: "Neighbour's garden", share: "35%" },
      { variety: "Cox's Orange Pippin", location: "Home orchard",   share: "25%" },
    ],
    gravity: 1.052,
    ph: 3.5,
    yeast: "Mangrove Jack's M02 Cider",
    volume: "23 L",
    readings: [
      { date: "2026-10-04", gravity: 1.052, ph: 3.5, note: "Pitched yeast" },
      { date: "2026-10-18", gravity: 1.020, note: "Steady fermentation" },
      { date: "2026-11-08", gravity: 1.004 },
    ],
    status: "Fermenting",
    notes: "Bittersweet / cooker blend. Racked once.",
  },
  "GL26-2": {
    vintage: 26,
    juiced: "2026-10-11",
    apples: [
      { variety: "Kingston Black", location: "Village orchard", share: "60%" },
      { variety: "Egremont Russet", location: "Home orchard",   share: "40%" },
    ],
    gravity: 1.058,
    ph: 3.4,
    yeast: "Lalvin EC-1118",
    readings: [
      { date: "2026-10-11", gravity: 1.058, ph: 3.4 },
      { date: "2026-11-01", gravity: 1.010 },
    ],
    status: "Fermenting",
  },
  "GL26-3": {
    vintage: 26,
    juiced: "2026-10-19",
    apples: [
      { variety: "Mixed windfalls", location: "Home orchard" },
      { variety: "Yarlington Mill", location: "Farm gate swap" },
    ],
    gravity: 1.048,
    ph: 3.7,
    yeast: "Wild (natural)",
    status: "Fermenting",
    notes: "Wild ferment - slow start.",
  },
};
