// -------------------------------------------------
// MSF ToolBot Import
// Author: chronolinq
// -------------------------------------------------

// Range Values
let tb_sheet_name = "*Toolbot";
let tb_id_cell = "A1";
let tb_roster_sheet_name = "Roster";
let tb_roster_sheet_range = "A2:X";
let r_data_range = "Roster_Data";

// GROOT Data Columns
let r_star = 17;
let r_redstar = 18;
let r_level = 19;
let r_gear = 20;
let r_gear_topleft = 21;
let r_gear_middleleft = 22;
let r_gear_bottomleft = 23;
let r_gear_topright = 24;
let r_gear_middleright = 25;
let r_gear_bottomright = 26;
let r_basic = 27;
let r_special = 28;
let r_ultimate = 29;
let r_passive = 30;
let r_power = 31;
let r_shards = 32;

// ToolBot Data Columns
let tb_id = 0;
let tb_power = 7;
let tb_star = 8;
let tb_redstar_active = 9;
let tb_redstar_inactive = 23;
let tb_level = 10;
let tb_gear = 11;
let tb_gear_topleft = 12;
let tb_gear_middleleft = 13;
let tb_gear_bottomleft = 14;
let tb_gear_topright = 15;
let tb_gear_middleright = 16;
let tb_gear_bottomright = 17;
let tb_basic = 18;
let tb_special = 19;
let tb_ultimate = 20;
let tb_passive = 21;
let tb_shards = 22;

// Keywords
let kw_clear_value = "None";
let kw_max_shards = "MAX";
let kw_equipped = "Y";

function msftoolbot_import()
{
    let toolbotSheetId;

    try
    {
        // Get Toolbot Sheet Id
        toolbotSheetId = SpreadsheetApp.getActive().getSheetByName(tb_sheet_name).getRange(tb_id_cell).getValue();
    }
    catch (err)
    {
        Browser.msgBox(
            "Cannot get MSF ToolBot Sheet ID",
            `Please ensure that a sheet named *Toolbot exists in this worksheet.\\n(It should appear as a tab on the bottom of the page)\\n\\nIf it exists, ensure that the MSF ToolBot Sheet ID in cell A1 is correct`,
            Browser.Buttons.OK);
        return;
    }

    let tbData;

    try
    {
        // Get ToolBot Range and Values
        let tbRange = SpreadsheetApp.openById(toolbotSheetId).getSheetByName(tb_roster_sheet_name).getRange(tb_roster_sheet_range);
        tbData = tbRange.getValues();
    }
    catch (err)
    {
        Browser.msgBox(
            "Cannot get MSF ToolBot data",
            `Please ensure the MSF ToolBot Sheet ID in the *Toolbot sheet is correct`,
            Browser.Buttons.OK);
        return;
    }

    // Get the data from the Roster sheet
    let rData = getNamedRangeValues(r_data_range);

    // Look through each row in the ToolBot range
    for (let i = 0; i < tbData.length; i++)
    {
        // Get the Id from the ToolBot row and try and get the data row
        let currentTbRow = tbData[i];
        let id = currentTbRow[tb_id];
        let currentData = getRowById(rData, id);

        // If this character is on the roster sheet, write out their info
        if (currentData)
        {
            setRosterValue(currentData, currentTbRow, r_level, tb_level);
            setRosterValue(currentData, currentTbRow, r_star, tb_star);
            setRosterRedstarValue(currentData, currentTbRow, r_redstar, tb_redstar_active, tb_redstar_inactive);
            setRosterValue(currentData, currentTbRow, r_power, tb_power);
            setRosterValue(currentData, currentTbRow, r_basic, tb_basic);
            setRosterValue(currentData, currentTbRow, r_special, tb_special);
            setRosterValue(currentData, currentTbRow, r_ultimate, tb_ultimate);
            setRosterValue(currentData, currentTbRow, r_passive, tb_passive);
            setRosterValue(currentData, currentTbRow, r_gear, tb_gear);
            setFarmingValue(currentData, currentTbRow, r_gear_topleft, tb_gear_topleft);
            setFarmingValue(currentData, currentTbRow, r_gear_middleleft, tb_gear_middleleft);
            setFarmingValue(currentData, currentTbRow, r_gear_bottomleft, tb_gear_bottomleft);
            setFarmingValue(currentData, currentTbRow, r_gear_topright, tb_gear_topright);
            setFarmingValue(currentData, currentTbRow, r_gear_middleright, tb_gear_middleright);
            setFarmingValue(currentData, currentTbRow, r_gear_bottomright, tb_gear_bottomright);
            setShardValue(currentData, currentTbRow, r_shards, tb_shards);
        }
    }

    // Set the updated data back on the ranges
    getNamedRange(r_data_range).setValues(rData);
}

// Set the ToolBot values on the Roster Organizer
function setRosterValue(rData, tbData, rIndex, tbIndex)
{
    let value = tbData[tbIndex];

    if (value == kw_clear_value) rData[rIndex] = "";
    else if (value) rData[rIndex] = value;
}

// Set the ToolBot redstar values on the Roster Organizer
function setRosterRedstarValue(rData, tbData, rIndex, tbRedstarActiveIndex, tbRedstarInactiveIndex)
{
    // The MSF ToolBot records active and inactive redstars in separate columns so they must be added
    let redstar = tbData[tbRedstarActiveIndex]
    let inactive = tbData[tbRedstarInactiveIndex]

    if (redstar == kw_clear_value) 
    {
        // Clear the value if Red Stars is "None"
        rData[rIndex] = "";

        // Set the red star value to the inactive value if it's not null and it's not "None"
        if (inactive && inactive != kw_clear_value)
        {
            rData[rIndex] = inactive;
        }

        return;
    }
    else if (redstar || inactive)
    {
        // If we have a value for inactive, and it's not "None", add it to Red Star
        if (inactive && inactive != kw_clear_value)
        {
            redstar += inactive;
        }

        rData[rIndex] = redstar;
    }
}

// Set the ToolBot shard values on the Roster Organizer
function setShardValue(rData, tbData, rIndex, toolbotIndex)
{
    // If the character's shards are maxed out, the value will be "MAX". The value should be cleared in that case
    let value = tbData[toolbotIndex];

    if (value == kw_clear_value || value == kw_max_shards) rData[rIndex] = "";
    else if (value) rData[rIndex] = value;
}

// Set the ToolBot equipped gear values on the Roster Organizer
function setFarmingValue(rData, tbData, rIndex, tbIndex)
{
    // If equipped, the value will be "Y", otherwise it will be "N"
    rData[rIndex] = (tbData[tbIndex] == kw_equipped) ? "TRUE" : "FALSE";
}

// Given a name, find the row that, that name exists
function getRowById(data, name)
{
    let row = null;

    for (let i = 0; i < data.length; i++)
    {
        if (data[i][0] == name)
        {
            row = data[i];
            break;
        }
    }

    return row;
}