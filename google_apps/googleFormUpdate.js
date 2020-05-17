// inspired by tutorial: http://wafflebytes.blogspot.com/2016/10/google-script-create-drop-down-list.html

function updateForm() {
    // call the form.
    let form = FormApp.openById(DMG_GEARING_PLAN_FORM_ID);
    
    // prepare map for form items.
    let formItemToOptionListMap = initFormItemToOptionListMap(form);
    
    // identify the sheet where the data resides to populate the drop-down
    let spreadSheet = SpreadsheetApp.getActive();
    
    // add all (key) items and (value) types to a map;
    let sheetItemToTypeMap = initSheetItemsToTypeMap(
        spreadSheet.getSheetByName("MC"),
        spreadSheet.getSheetByName("ONY"),
        spreadSheet.getSheetByName("BWL"),
        spreadSheet.getSheetByName("WB")
    );
    
    // populate form drop down menus with item data
    addMapItemsToForm(sheetItemToTypeMap, formItemToOptionListMap);
}

function addMapItemsToForm(sheetItemToTypeMap, formItemToOptionListMap) {
    for(formMenuTypeKey of FORM_MENU_TYPE_TO_ID_MAP.keys()) {
        addMapItemsByKeyToForm(sheetItemToTypeMap, formItemToOptionListMap, formMenuTypeKey)
    }
}

function addMapItemsByKeyToForm(sheetItemToTypeMap, formItemToOptionListMap, formMenuTypeKey) {
    // prepare arrays for possible menu options
    let options = [];
    let ringOptions = [];
    let trinketOptions = [];
    let weaponOptions = [];
    // Check key for a formKey match, or match with item types that have multiple slot possibilities. (Ring, Trinket, and Weapon)
    for(let [key,value] of sheetItemToTypeMap){
        let strKey = key.valueOf().toString();
        let strValue = value.valueOf().toString();
        
        if(strValue.includes('Ring')){
            ringOptions.push(strKey); // shows in both Ring1 and Ring2 menus.
        } else if (strValue.includes('Trinket')){
            trinketOptions.push(strKey); // shows in both Trinket1 and Trinket2 menus.
        } else if (strValue.includes('Weapon')){
            weaponOptions.push(strKey); // shows in both MainHand and OffHand menus.
        } else if (strValue === formMenuTypeKey){
            options.push(strKey); // remaining formKeys.
        }
    }
    
    // Sort options alphabetically A-Z for menus.
    // push all options to correct form menu by associated form key
    if(formMenuTypeKey == RING_1 || formMenuTypeKey == RING_2){
        formItemToOptionListMap.get(formMenuTypeKey).setChoiceValues(ringOptions.sort());
    }
    else if (formMenuTypeKey == TRINKET_1 || formMenuTypeKey == TRINKET_2){
        trinketOptions.sort();
        formItemToOptionListMap.get(formMenuTypeKey).setChoiceValues(trinketOptions);
    }
    else if (formMenuTypeKey == MAIN_HAND || formMenuTypeKey == OFF_HAND){
        weaponOptions.sort();
        formItemToOptionListMap.get(formMenuTypeKey).setChoiceValues(weaponOptions);
    }
    else {
        options.sort();
        formItemToOptionListMap.get(formMenuTypeKey).setChoiceValues(options);
    }
}

function initSheetItemsToTypeMap(...sheets) {
    let sheetItemsToItemTypeMap = new Map();
    
    for(sheet of sheets){
        // grab the values in the first and second column of the sheet - use 2 to skip header row.
        let itemColumnData = getCellData(2, 1, sheet);
        let itemTypeColumnData = getCellData(2, 2, sheet);
    
        // convert an array ignoring empty cells.
        let itemsArray = [];
        let itemTypesArray = [];
        
        for(let i = 0; i < itemColumnData.length; i++) {
            if(itemColumnData[i][0] != "") {
                itemsArray[i] = itemColumnData[i];
                itemTypesArray[i] = itemTypeColumnData[i];
            }
        }
    
        for(let i = 0; i< itemsArray.length; i++){
            sheetItemsToItemTypeMap.set(itemsArray[i], itemTypesArray[i]);
        }
    }
    
    return sheetItemsToItemTypeMap;
}

function getCellData(startRow, column, sheet) {
    return sheet.getRange(startRow, column, sheet.getMaxRows() - 1).getValues();
}

function getFormItemAsList(form, itemId) {
    return form.getItemById(itemId).asListItem()
}

function initFormItemToOptionListMap(form) {
    let formItemToOptionListMap = new Map();
    for(let [key,value] of FORM_MENU_TYPE_TO_ID_MAP){
        formItemToOptionListMap.set(key, getFormItemAsList(form, value));
    }
    return formItemToOptionListMap;
}