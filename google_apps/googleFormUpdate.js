// inspired by tutorial: http://wafflebytes.blogspot.com/2016/10/google-script-create-drop-down-list.html

function updateForm() {
    // call the form.
    let form = FormApp.openById("1HPfYs1z4lUkaDWF1LGxS6PQCnpvjbp6AHx35UQT83nc");
    
    // prepare map for form items.
    let formItemListMap = new Map();
    initFormMap(form, formItemListMap)
    
    // identify the sheet where the data resides to populate the drop-down
    let ss = SpreadsheetApp.getActive();
    
    // get relevant sheet data.
    let sheetMC = ss.getSheetByName("MC");
    let sheetONY = ss.getSheetByName("ONY");
    let sheetBWL = ss.getSheetByName("BWL");
    let sheetZG = ss.getSheetByName("ZG");
    let sheetWB = ss.getSheetByName("WB");
    
    // add all (key) items and (value) type to the ItemToTypeMap;
    let sheetItemToTypeMap = new Map();
    addSheetItemsToMap(sheetMC, sheetItemToTypeMap);
    addSheetItemsToMap(sheetONY, sheetItemToTypeMap);
    addSheetItemsToMap(sheetBWL, sheetItemToTypeMap);
    addSheetItemsToMap(sheetZG, sheetItemToTypeMap);
    addSheetItemsToMap(sheetWB, sheetItemToTypeMap);
    
    // // populate the drop-down with the array data
    addMapItemsToForm(sheetItemToTypeMap, formItemListMap);
}

function addMapItemsToForm(itemToTypeMap, formMap) {
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Head");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Neck");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Shoulder");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Back");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Chest");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Wrist");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Hands");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Waist");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Legs");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Feet");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Ring1");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Ring2");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Trinket1");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Trinket2");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "MainHand");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "OffHand");
    addMapItemsByKeyToForm(itemToTypeMap, formMap, "Ranged");
}

function addMapItemsByKeyToForm(itemToTypeMap, formMap, formKey) {
    let options = [];
    let ringOptions = [];
    let trinketOptions = [];
    let weaponOptions = [];
    
    for(let [key,value] of itemToTypeMap){
        let strKey = key.valueOf().toString();
        let strValue = value.valueOf().toString();
        
        if(strValue.includes('Ring')){
            ringOptions.push(strKey);
        } else if (strValue.includes('Trinket')){
            trinketOptions.push(strKey);
        } else if (strValue.includes('Weapon')){
            weaponOptions.push(strKey);
        } else if (strValue === formKey){
            options.push(strKey);
        }
    }
    
    if(formKey == "Ring1" || formKey == "Ring2"){
        formMap.get(formKey).setChoiceValues(ringOptions);
    } else if (formKey == "Trinket1" || formKey == "Trinket2"){
        formMap.get(formKey).setChoiceValues(trinketOptions);
    } else if (formKey == "MainHand" || formKey == "OffHand"){
        formMap.get(formKey).setChoiceValues(weaponOptions);
    } else {
        formMap.get(formKey).setChoiceValues(options);
    }
}

function addSheetItemsToMap(sheet, map) {
    // grab the values in the first and second column of the sheet - use 2 to skip header row.
    let col1 = getCellData(2, 1, sheet);
    let col2 = getCellData(2, 2, sheet);
    
    // convert an array ignoring empty cells.
    let items = []
    setOptions(col1, items);
    let itemTypes = []
    setOptions(col2, itemTypes);
    
    for(let i = 0; i< items.length; i++){
        map.set(items[i], itemTypes[i]);
    }
}

function getCellData(startRow, column, sheet) {
    return sheet.getRange(startRow, column, sheet.getMaxRows() - 1).getValues();
}

function setOptions(values, arrayToFill) {
    for(let i = 0; i < values.length; i++) {
        if(values[i][0] != "") {
            arrayToFill[i] = values[i];
        }
    }
}

function getFormItemAsList(form, itemId) {
    return form.getItemById(itemId).asListItem()
}

function initFormMap(form, map) {
    map.set("Head", getFormItemAsList(form, "407218439"));
    map.set("Neck", getFormItemAsList(form, "1793300395"));
    map.set("Shoulder", getFormItemAsList(form, "1134373307"));
    map.set("Back", getFormItemAsList(form, "1922841249"));
    map.set("Chest", getFormItemAsList(form, "98938550"));
    map.set("Wrist", getFormItemAsList(form, "453815564"));
    map.set("Hands", getFormItemAsList(form, "603291189"));
    map.set("Waist", getFormItemAsList(form, "704089010"));
    map.set("Legs", getFormItemAsList(form, "489504314"));
    map.set("Feet", getFormItemAsList(form, "600582489"));
    map.set("Ring1", getFormItemAsList(form, "1796460282"));
    map.set("Ring2", getFormItemAsList(form, "1018155902"));
    map.set("Trinket1", getFormItemAsList(form, "202851639"));
    map.set("Trinket2", getFormItemAsList(form, "977531088"));
    map.set("MainHand", getFormItemAsList(form, "1112911872"));
    map.set("OffHand", getFormItemAsList(form, "396892926"));
    map.set("Ranged", getFormItemAsList(form, "1844595494"));
}