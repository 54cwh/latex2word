let imageDir = null;
let formulaTracker = null;
let footnoteCounter = 0;
let footnoteTexts = {};
let citationMap = {};
let eqCounter = 0;
let sectionNumber = 1;

function setFormulaTracker(tracker) {
  formulaTracker = tracker;
}

function saveState() {
  return { formulaTracker, footnoteCounter, footnoteTexts, citationMap, eqCounter, sectionNumber, imageDir };
}

function restoreState(s) {
  if (s) { formulaTracker = s.formulaTracker; footnoteCounter = s.footnoteCounter; footnoteTexts = s.footnoteTexts; citationMap = s.citationMap; eqCounter = s.eqCounter; sectionNumber = s.sectionNumber; imageDir = s.imageDir; }
}

function resetState() {
  formulaTracker = null;
  footnoteCounter = 0;
  footnoteTexts = {};
  citationMap = {};
  eqCounter = 0;
  sectionNumber = 1;
  imageDir = null;
}

module.exports = {
  get imageDir() { return imageDir; },
  set imageDir(v) { imageDir = v; },
  get formulaTracker() { return formulaTracker; },
  set formulaTracker(v) { formulaTracker = v; },
  get footnoteCounter() { return footnoteCounter; },
  set footnoteCounter(v) { footnoteCounter = v; },
  get footnoteTexts() { return footnoteTexts; },
  set footnoteTexts(v) { footnoteTexts = v; },
  get citationMap() { return citationMap; },
  set citationMap(v) { citationMap = v; },
  get eqCounter() { return eqCounter; },
  set eqCounter(v) { eqCounter = v; },
  get sectionNumber() { return sectionNumber; },
  set sectionNumber(v) { sectionNumber = v; },
  setFormulaTracker, saveState, restoreState, resetState
};
