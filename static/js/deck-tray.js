// Shared deck-tray (localStorage) state used by the build and library pages.
const DECK_KEY='j2w_deck';
function loadDeck(){try{return JSON.parse(localStorage.getItem(DECK_KEY));}catch(e){return null;}}
function saveDeck(d){localStorage.setItem(DECK_KEY,JSON.stringify(d));}
