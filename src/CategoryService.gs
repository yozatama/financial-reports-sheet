/**
 * CategoryService.gs
 * -------------------------------------------------------------
 * Lightweight CRUD for the Categories sheet plus helpers used by
 * the AI parser to map free-text words to a known category.
 * -------------------------------------------------------------
 */

var CategoryService = (function () {

  function add(c) {
    require_(c, 'Category payload required');
    require_(c.type, 'Type required');
    require_(c.category, 'Category required');
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CATEGORIES);
    var row = [
      c.id || genId('CAT'),
      c.type, c.category, c.subcategory || '',
      c.icon || '', c.color || '#9CA3AF', c.active === false ? 'No' : 'Yes'
    ];
    sh.appendRow(row);
    return row;
  }

  function update(id, patch) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CATEGORIES);
    var rowIdx = findRowByValue_(sh, COLS.CATEGORIES.ID, id);
    if (rowIdx < 0) throw new Error('Category not found');
    var row = sh.getRange(rowIdx, 1, 1, 7).getValues()[0];
    if (patch.type) row[COLS.CATEGORIES.TYPE - 1] = patch.type;
    if (patch.category) row[COLS.CATEGORIES.CATEGORY - 1] = patch.category;
    if (patch.subcategory !== undefined) row[COLS.CATEGORIES.SUBCATEGORY - 1] = patch.subcategory;
    if (patch.icon !== undefined) row[COLS.CATEGORIES.ICON - 1] = patch.icon;
    if (patch.color !== undefined) row[COLS.CATEGORIES.COLOR - 1] = patch.color;
    if (patch.active !== undefined) row[COLS.CATEGORIES.ACTIVE - 1] = patch.active ? 'Yes' : 'No';
    sh.getRange(rowIdx, 1, 1, 7).setValues([row]);
    return row;
  }

  function remove(id) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CATEGORIES);
    var rowIdx = findRowByValue_(sh, COLS.CATEGORIES.ID, id);
    if (rowIdx < 0) throw new Error('Category not found');
    sh.deleteRow(rowIdx);
    return { ok: true };
  }

  function list(type) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CATEGORIES);
    if (!sh) return [];
    var rows = readObjects_(sh).filter(function (r) { return r.Active !== 'No'; });
    return type ? rows.filter(function (r) { return r.Type === type; }) : rows;
  }

  function topLevelNames(type) {
    return unique_(list(type).map(function (r) { return r.Category; }));
  }

  function exists(name) {
    return list().some(function (r) { return String(r.Category).toLowerCase() === String(name).toLowerCase(); });
  }

  return {
    add: add, update: update, remove: remove,
    list: list, topLevelNames: topLevelNames, exists: exists
  };
})();
