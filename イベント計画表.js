/**
 * onEdit - 改訂版（入力日付が見つからない場合は編集セルにノートを付与）
 * - E列(5)=計画開始日 -> '始' / '始/期' / '⇒'
 * - F列(6)=期日日程   -> '期' / '始/期' / '⇒'
 * - G列(7)=完了日     -> '済' （Gクリア時に必ず E -> F の順で再処理）
 * - 色は使わない
 * - 日付が H列(3行目) に見つからないときは編集セルに note を付与
 */

function onEdit(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) return;
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var row = range.getRow();
    var col = range.getColumn();
    var value = range.getValue();

    if (row < 4) return;

    var tz = Session.getScriptTimeZone();

    function fmt(v) {
      if (v instanceof Date) return Utilities.formatDate(v, tz, "MM/dd");
      if (typeof v === 'string') return v.trim();
      if (v === null || v === undefined) return '';
      return String(v).trim();
    }

    function findDateColumns(sheet, targetFormatted) {
      var startCol = 8; // H
      var lastCol = sheet.getLastColumn();
      if (lastCol < startCol) return [];
      var headerRow = sheet.getRange(3, startCol, 1, lastCol - startCol + 1).getValues()[0];
      var matches = [];
      for (var i = 0; i < headerRow.length; i++) {
        var hv = headerRow[i];
        var hf = fmt(hv);
        if (hf === targetFormatted) matches.push(startCol + i);
      }
      return matches;
    }

    function scanRowStates(sheet, row) {
      var startCol = 8;
      var lastCol = sheet.getLastColumn();
      var res = { starts: [], terms: [], boths: [], dones: [], arrows: [] };
      if (lastCol < startCol) return res;
      var vals = sheet.getRange(row, startCol, 1, lastCol - startCol + 1).getValues()[0];
      for (var i = 0; i < vals.length; i++) {
        var v = (vals[i] === null) ? '' : String(vals[i]).trim();
        var c = startCol + i;
        if (v === '始') res.starts.push(c);
        else if (v === '期') res.terms.push(c);
        else if (v === '始/期' || v === '期/始') res.boths.push(c);
        else if (v === '済') res.dones.push(c);
        else if (v === '⇒') res.arrows.push(c);
      }
      return res;
    }

    function clearArrowsBetween(sheet, row, c1, c2) {
      if (c2 < c1) return;
      var rng = sheet.getRange(row, c1, 1, c2 - c1 + 1);
      var vals = rng.getValues()[0];
      var changed = false;
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i]).trim() === '⇒') { vals[i] = ''; changed = true; }
      }
      if (changed) rng.setValues([vals]);
    }

    function clearAllArrows(sheet, row) {
      var startCol = 8;
      var lastCol = sheet.getLastColumn();
      if (lastCol < startCol) return;
      var rng = sheet.getRange(row, startCol, 1, lastCol - startCol + 1);
      var vals = rng.getValues()[0];
      var changed = false;
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i]).trim() === '⇒') { vals[i] = ''; changed = true; }
      }
      if (changed) rng.setValues([vals]);
    }

    function setArrowRange(sheet, row, from, to) {
      if (to < from) return;
      var len = to - from + 1;
      var arr = [];
      for (var i = 0; i < len; i++) arr.push('⇒');
      sheet.getRange(row, from, 1, len).setValues([arr]);
    }

    function setCell(sheet, row, col, val) {
      sheet.getRange(row, col).setValue(val);
    }

    function updateArrowsForRow(sheet, row) {
      var st = scanRowStates(sheet, row);
      clearAllArrows(sheet, row);
      if (st.starts.length > 0 && st.terms.length > 0) {
        var s = st.starts[0], t = st.terms[0];
        var from = Math.min(s, t) + 1; // exclude endpoints
        var to = Math.max(s, t) - 1;
        if (from <= to) setArrowRange(sheet, row, from, to);
      }
    }

    // ---- processing functions now accept inputCol to manage notes ----
    function processStart(sheet, row, formattedValue, inputCol) {
      sheet.getRange(row, inputCol).clearNote();
      var matchCols = findDateColumns(sheet, formattedValue);
      if (!matchCols || matchCols.length === 0) {
        sheet.getRange(row, inputCol).setNote("3行目のH列以降に '" + formattedValue + "' が見つかりません。");
        return;
      }
      var st = scanRowStates(sheet, row);

      if (st.boths.length > 0) {
        st.boths.forEach(function(bc){ setCell(sheet, row, bc, '期'); });
        st = scanRowStates(sheet, row);
      }

      if (st.starts.length > 0) {
        var oldStart = st.starts[0];
        if (st.terms.length > 0) {
          var oldTerm = st.terms[0];
          var aFrom = Math.min(oldStart, oldTerm) + 1;
          var aTo = Math.max(oldStart, oldTerm) - 1;
          clearArrowsBetween(sheet, row, aFrom, aTo);
        } else {
          clearAllArrows(sheet, row);
        }
        sheet.getRange(row, oldStart).clearContent();
        st = scanRowStates(sheet, row);
      }

      matchCols.forEach(function(mc) {
        if (st.terms.indexOf(mc) !== -1) {
          setCell(sheet, row, mc, '始/期');
          clearAllArrows(sheet, row);
          return;
        }
        if (st.terms.length > 0) {
          var termCol = st.terms[0];
          setCell(sheet, row, mc, '始');
          var aFrom = Math.min(mc, termCol) + 1;
          var aTo = Math.max(mc, termCol) - 1;
          clearAllArrows(sheet, row);
          if (aFrom <= aTo) setArrowRange(sheet, row, aFrom, aTo);
          return;
        }
        setCell(sheet, row, mc, '始');
      });
    }

    function processTerm(sheet, row, formattedValue, inputCol) {
      sheet.getRange(row, inputCol).clearNote();
      var matchCols = findDateColumns(sheet, formattedValue);
      if (!matchCols || matchCols.length === 0) {
        sheet.getRange(row, inputCol).setNote("3行目のH列以降に '" + formattedValue + "' が見つかりません。");
        return;
      }
      var st = scanRowStates(sheet, row);

      if (st.boths.length > 0) {
        st.boths.forEach(function(bc){ setCell(sheet, row, bc, '始'); });
        st = scanRowStates(sheet, row);
      }

      if (st.terms.length > 0) {
        var oldTerm = st.terms[0];
        if (st.starts.length > 0) {
          var oldStart = st.starts[0];
          var aFrom = Math.min(oldStart, oldTerm) + 1;
          var aTo = Math.max(oldStart, oldTerm) - 1;
          clearArrowsBetween(sheet, row, aFrom, aTo);
        } else {
          clearAllArrows(sheet, row);
        }
        sheet.getRange(row, oldTerm).clearContent();
        st = scanRowStates(sheet, row);
      }

      matchCols.forEach(function(mc) {
        if (st.starts.indexOf(mc) !== -1) {
          setCell(sheet, row, mc, '始/期');
          clearAllArrows(sheet, row);
          return;
        }
        setCell(sheet, row, mc, '期');
        if (st.starts.length > 0) {
          var startCol = st.starts[0];
          var aFrom = Math.min(startCol, mc) + 1;
          var aTo = Math.max(startCol, mc) - 1;
          clearAllArrows(sheet, row);
          if (aFrom <= aTo) setArrowRange(sheet, row, aFrom, aTo);
        }
      });
    }

    function processDone(sheet, row, formattedValue, inputCol) {
      sheet.getRange(row, inputCol).clearNote();
      var matchCols = findDateColumns(sheet, formattedValue);
      if (!matchCols || matchCols.length === 0) {
        sheet.getRange(row, inputCol).setNote("3行目のH列以降に '" + formattedValue + "' が見つかりません。");
        return;
      }
      var st = scanRowStates(sheet, row);

      if (st.dones.length > 0) {
        st.dones.forEach(function(dc){ sheet.getRange(row, dc).clearContent(); });
        st = scanRowStates(sheet, row);
      }

      matchCols.forEach(function(mc){ setCell(sheet, row, mc, '済'); });
    }

    // handleClear: E/F/G それぞれのマーカーだけクリア（Gクリア時は必ず E->F の順で再処理）
    function handleClear(sheet, row, clearedCol) {
      var st = scanRowStates(sheet, row);

      // clear note on the cleared input cell
      sheet.getRange(row, clearedCol).clearNote();

      if (clearedCol === 5) { // E cleared
        if (st.boths.length > 0) {
          st.boths.forEach(function(bc){ setCell(sheet, row, bc, '期'); });
          updateArrowsForRow(sheet, row);
          return;
        } else {
          if (st.starts.length > 0) {
            st.starts.forEach(function(c) {
              if (st.terms.length > 0) {
                var term = st.terms[0];
                var aFrom = Math.min(c, term) + 1;
                var aTo = Math.max(c, term) - 1;
                if (aFrom <= aTo) clearArrowsBetween(sheet, row, aFrom, aTo);
              } else {
                clearAllArrows(sheet, row);
              }
              sheet.getRange(row, c).clearContent();
            });
          }
          return;
        }
      } else if (clearedCol === 6) { // F cleared
        if (st.boths.length > 0) {
          st.boths.forEach(function(bc){ setCell(sheet, row, bc, '始'); });
          updateArrowsForRow(sheet, row);
          return;
        } else {
          if (st.terms.length > 0) {
            st.terms.forEach(function(c) {
              if (st.starts.length > 0) {
                var start = st.starts[0];
                var aFrom = Math.min(start, c) + 1;
                var aTo = Math.max(start, c) - 1;
                if (aFrom <= aTo) clearArrowsBetween(sheet, row, aFrom, aTo);
              } else {
                clearAllArrows(sheet, row);
              }
              sheet.getRange(row, c).clearContent();
            });
          }
          return;
        }
      } else if (clearedCol === 7) { // G cleared
        // clear only '済'
        if (st.dones.length > 0) {
          st.dones.forEach(function(c){ sheet.getRange(row, c).clearContent(); });
        }
        // reprocess E -> F in that order (guaranteed)
        var eVal = fmt(sheet.getRange(row, 5).getValue());
        var fVal = fmt(sheet.getRange(row, 6).getValue());
        if (eVal !== '') processStart(sheet, row, eVal, 5);
        // re-scan is implicit inside processTerm; still call processTerm after E
        if (fVal !== '') processTerm(sheet, row, fVal, 6);
        return;
      }
    }

    // dispatch
    if (col === 5) { // E
      if (fmt(value) === '') handleClear(sheet, row, 5);
      else processStart(sheet, row, fmt(value), 5);
    } else if (col === 6) { // F
      if (fmt(value) === '') handleClear(sheet, row, 6);
      else processTerm(sheet, row, fmt(value), 6);
    } else if (col === 7) { // G
      if (fmt(value) === '') handleClear(sheet, row, 7);
      else processDone(sheet, row, fmt(value), 7);
    } else {
      // nothing
    }

  } catch (err) {
    Logger.log('onEdit error: ' + err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
