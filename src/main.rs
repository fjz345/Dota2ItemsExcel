extern crate simple_excel_writer as excel;
use excel::*;
use scraper::{Html, Selector};

use serde_json::{Map, Value};
use std::process::Command;

#[derive(Debug)]
struct ItemEntry
{
    Name: String,
    Category: String
}

#[derive(Debug)]
#[derive(Clone)]
struct Hero
{
    Name: String,
    PrimaryAttribute: String,
    AttackType: String,
    BAT: f32,
    BaseAttackSpeed: i32
}

#[derive(Debug)]
#[derive(Clone)]
struct Item
{
    Name: String,
    Damage: i32,
    Damage_Melee: i32,
    Damage_Ranged: i32,
    AttackSpeed: i32,
    Str: i32,
    Agi: i32,
    Int: i32,
    ArmorCorruption: i32,
    MagicDamage: i32,
    MagicChance_Melee: f32,
    MagicChance_Ranged: f32,
    CritMultiplier: f32,
    CritChance: f32,
    Cost: i32,
    IsNeutralItem: bool,
    IsUselessItem: bool,
}

impl Default for Item {
    fn default() -> Item {
        Item {
            Name: "Unset".to_string(),
            Damage: 0,
            Damage_Melee: 0,
            Damage_Ranged: 0,
            AttackSpeed: 0,
            ArmorCorruption: 0,
            MagicDamage: 0,
            MagicChance_Melee: 0.0,
            MagicChance_Ranged: 0.0,
            CritMultiplier: 1.0,
            CritChance: 0.0,
            Str: 0,
            Agi: 0,
            Int: 0,
            Cost: 0,
            IsNeutralItem: false,
            IsUselessItem: true,
        }
    }
}

fn main() {

    // Item list
    let ItemDataJson = GetItemDataJsonString();
    let mut Items: Vec<Item> = Vec::new();

    GetItemStats(&ItemDataJson, &mut Items, true);

    // Replace item_names with real names
    GetRealItemNames(&mut Items);

    let mut wb = Workbook::create("../Dota2Data.xlsx");
    WriteItemsToXlsx(&mut wb, &Items);




    // Hero list
    let HeroDataJson = GetHeroDataJsonString();
    let mut HeroList: Vec<Hero> = Vec::new();

    GetHeroesData(&HeroDataJson, &mut HeroList);

    WriteHeroesToXlsx(&mut wb, &HeroList);

    // Close
    wb.close().expect("close excel error!");

    // Open the file
    let output = if cfg!(target_os = "windows") {
        Command::new("cmd")
                .args(["/C", "start ../Dota2BuyDps.xlsm"])
                .output()
                .expect("failed to execute process")
    } else {
        Command::new("sh")
                .arg("-c")
                .arg("../Dota2BuyDps.xlsm")
                .output()
                .expect("failed to execute process")
    };
}

fn GetItemDataJsonString() -> String
{
    // Get html source
    let url = "https://raw.githubusercontent.com/dotabuff/d2vpkr/master/dota/scripts/npc/items.json";
    let response = reqwest::blocking::get(url).unwrap();
    let htmlBody = response.text().unwrap();
    htmlBody
}

fn GetHeroDataJsonString() -> String
{
    // Get html source
    let url = "https://raw.githubusercontent.com/odota/dotaconstants/master/build/heroes.json";
    let response = reqwest::blocking::get(url).unwrap();
    let htmlBody = response.text().unwrap();
    htmlBody
}

fn GetRealItemNames(InOutItems: &mut Vec<Item>)
{
    // Get html source
    let url = "https://raw.githubusercontent.com/odota/dotaconstants/master/build/items.json";
    let response = reqwest::blocking::get(url).unwrap();
    let json = response.text().unwrap();

    let allItemsDotaConstants: Map<String, Value> = serde_json::from_str(&json).unwrap();

    for item in InOutItems
    {
        // remove item_
        let str = &item.Name[5..];

        if(allItemsDotaConstants.contains_key(str))
        {
            let itemInfoMap: Map<String, Value> = serde_json::from_value(allItemsDotaConstants[str].clone()).unwrap();

            if(itemInfoMap.contains_key("dname"))
            {
                item.Name = itemInfoMap["dname"].as_str().unwrap().to_string();
            }
        }
    }
}

fn GetHeroesData(JsonData: &String, InOutHeroes: &mut Vec<Hero>) 
{
    let parsed: Map<String, Value> = serde_json::from_str(JsonData).unwrap();
    let allHeroes = parsed.clone();

    // Loop through all items, if matches criterias, add it to the item list
    for hero in &allHeroes
    {
        //println!("{:#?}", hero.1);

        let heroMap: Map<String, Value> = serde_json::from_value(hero.1.clone()).unwrap();

        let aName = heroMap["localized_name"].as_str().unwrap().to_string();
        let aPrimaryAttribute = heroMap["primary_attr"].as_str().unwrap().to_string();
        let aAttackType = heroMap["attack_type"].as_str().unwrap().to_string();
        let aBAT = heroMap["attack_rate"].as_f64().unwrap().to_string().parse::<f32>().unwrap();
        let aBaseAttackSpeed = 100;
        let mut aHero: Hero = Hero{Name: aName, PrimaryAttribute: aPrimaryAttribute, AttackType: aAttackType, BAT: aBAT, BaseAttackSpeed: aBaseAttackSpeed};

        InOutHeroes.push(aHero);
    }
}

fn GetItemStats(JsonData: &String, InOutItems: &mut Vec<Item>, IgnoreUselessItems: bool) 
{
    let parsed: Map<String, Value> = serde_json::from_str(JsonData).unwrap();
    let mut versionIncluded: Map<String, Value> = serde_json::from_str(&parsed["DOTAAbilities"].to_string()).unwrap();
    versionIncluded.remove("Version");
    let allItems = versionIncluded.clone();

    // Loop through all items, if matches criterias, add it to the item list
    for item in &allItems
    {
        //println!("{:#?}", item.0);

        let mut isUselessItem = true;

        let aName = item.0.clone();
        let mut aItem: Item = Item{Name: aName, ..Default::default()};
        
        let itemMap: Map<String, Value> = serde_json::from_value(item.1.clone()).unwrap();

        // Early Continue
        if itemMap.contains_key("IsObsolete")
        {
            if itemMap["IsObsolete"].as_str().unwrap().parse::<i32>().unwrap() > 0
            {
                continue;
            }
        }

        // Fill all values from json

        // Cost
        if (itemMap.contains_key("ItemCost"))
        {
            // Special case
            if(itemMap["ItemCost"].as_str().unwrap().to_string() == "")
            {
                aItem.Cost = 0;
            }
            else
            {
                aItem.Cost = itemMap["ItemCost"].as_str().unwrap().to_string().parse::<i32>().unwrap();
            }
        }

        // IsNeutralItem
        if (itemMap.contains_key("ItemIsNeutralDrop"))
        {
            // Special case
            if(itemMap["ItemIsNeutralDrop"].as_str().unwrap().to_string() == "")
            {
                aItem.IsNeutralItem = false;
            }
            else if (itemMap["ItemIsNeutralDrop"].as_str().unwrap().to_string() == "0")
            {
                aItem.IsNeutralItem = false;
            }
            else if (itemMap["ItemIsNeutralDrop"].as_str().unwrap().to_string() == "1")
            {
                aItem.IsNeutralItem = true;
            }
        }

        if (itemMap.contains_key("AbilitySpecial"))
        {
            let bonusMap: Vec<Map<String,Value>> = serde_json::from_value(itemMap["AbilitySpecial"].clone()).unwrap();

            for attribute in &bonusMap[..]
            {
                // Str
                if (attribute.contains_key("bonus_strength"))
                {
                    if(attribute["bonus_strength"].is_number())
                    {
                        aItem.Str = attribute["bonus_strength"].as_i64().unwrap() as i32;
                        isUselessItem = false;
                    }
                    else if(attribute["bonus_strength"].is_string())
                    {
                        aItem.Str = attribute["bonus_strength"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                        isUselessItem = false;
                    }
                }

                // Agi
                else if (attribute.contains_key("bonus_agility"))
                {
                    if(attribute["bonus_agility"].is_number())
                    {
                        aItem.Agi = attribute["bonus_agility"].as_i64().unwrap() as i32;
                        isUselessItem = false;
                    }
                    else if(attribute["bonus_agility"].is_string())
                    {
                        aItem.Agi = attribute["bonus_agility"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                        isUselessItem = false;
                    }
                }

                // Int
                else if (attribute.contains_key("bonus_intellect"))
                {
                    if(attribute["bonus_intellect"].is_number())
                    {
                        aItem.Int = attribute["bonus_intellect"].as_i64().unwrap() as i32;
                        isUselessItem = false;
                    }
                    else if(attribute["bonus_intellect"].is_string())
                    {
                        aItem.Int = attribute["bonus_intellect"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                        isUselessItem = false;
                    }
                }

                // Str,Agi,Int
                else if (attribute.contains_key("bonus_all_stats"))
                {
                    if(attribute["bonus_all_stats"].is_number())
                    {
                        if(attribute["bonus_all_stats"].as_i64().unwrap() as i32 != 0)
                        {
                            aItem.Str += attribute["bonus_all_stats"].as_i64().unwrap() as i32;
                            aItem.Agi += attribute["bonus_all_stats"].as_i64().unwrap() as i32;
                            aItem.Int += attribute["bonus_all_stats"].as_i64().unwrap() as i32;
                            isUselessItem = false;
                        }
                    }
                    else if(attribute["bonus_all_stats"].is_string())
                    {
                        if(attribute["bonus_all_stats"].as_str().unwrap().to_string().parse::<i32>().unwrap() != 0)
                        {
                            aItem.Str += attribute["bonus_all_stats"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                            aItem.Agi += attribute["bonus_all_stats"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                            aItem.Int += attribute["bonus_all_stats"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                            isUselessItem = false;
                        }
                    }
                }

                // Damage
                else if (attribute.contains_key("bonus_damage"))
                {
                    // Enchanted quiver damage is on cd
                    if(item.0 != "item_enchanted_quiver")
                    {
                        if(attribute["bonus_damage"].is_number())
                        {
                            if(attribute["bonus_damage"].is_f64())
                            {
                                aItem.Damage = attribute["bonus_damage"].as_f64().unwrap() as i32;
                                isUselessItem = false;
                            }
                            else
                            {
                                aItem.Damage = attribute["bonus_damage"].as_i64().unwrap() as i32;
                                isUselessItem = false;
                            }
                        }
                        else if(attribute["bonus_damage"].is_string())
                        {
                            aItem.Damage = attribute["bonus_damage"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                            isUselessItem = false;
                        }
                    }
                }

                // Damage Melee
                else if (attribute.contains_key("bonus_damage_melee"))
                {
                    if(attribute["bonus_damage_melee"].is_number())
                    {
                        aItem.Damage_Melee = attribute["bonus_damage_melee"].as_i64().unwrap() as i32;
                        isUselessItem = false;
                    }
                    else if(attribute["bonus_damage_melee"].is_string())
                    {
                        aItem.Damage_Melee = attribute["bonus_damage_melee"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                        isUselessItem = false;
                    }
                }

                // Damage Ranged
                else if (attribute.contains_key("bonus_damage_range"))
                {
                    if(attribute["bonus_damage_range"].is_number())
                    {
                        aItem.Damage_Ranged = attribute["bonus_damage_range"].as_i64().unwrap() as i32;
                        isUselessItem = false;
                    }
                    else if(attribute["bonus_damage_range"].is_string())
                    {
                        aItem.Damage_Ranged = attribute["bonus_damage_range"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                        isUselessItem = false;
                    }
                }

                // Attack Speed
                else if (attribute.contains_key("bonus_attack_speed"))
                {
                    // Hurricanes attack speed is activate
                    if(item.0 != "item_hurricane_pike")
                    {
                        if(attribute["bonus_attack_speed"].is_number())
                        {
                            aItem.AttackSpeed = attribute["bonus_attack_speed"].as_i64().unwrap() as i32;
                            isUselessItem = false;
                        }
                        else if(attribute["bonus_attack_speed"].is_string())
                        {
                            aItem.AttackSpeed = attribute["bonus_attack_speed"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                            isUselessItem = false;
                        }
                    }
                }

                // Armor Corr
                else if (attribute.contains_key("corruption_armor"))
                {
                    if(attribute["corruption_armor"].is_number())
                    {
                        aItem.ArmorCorruption = attribute["corruption_armor"].as_i64().unwrap() as i32;
                        isUselessItem = false;
                    }
                    else if(attribute["corruption_armor"].is_string())
                    {
                        aItem.ArmorCorruption = attribute["corruption_armor"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                        isUselessItem = false;
                    }
                }

                // For some reason source json has "armor" as corruption for item_orb_of_corrosion
                else if (attribute.contains_key("armor"))
                {
                    if(item.0 == "item_orb_of_corrosion")
                    {
                        if(attribute["armor"].is_number())
                        {
                            aItem.ArmorCorruption = -attribute["armor"].as_i64().unwrap() as i32;
                            isUselessItem = false;
                        }
                        else if(attribute["armor"].is_string())
                        {
                            aItem.ArmorCorruption = -attribute["armor"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                            isUselessItem = false;
                        }
                    }
                }

                // Magic Damage
                else if (attribute.contains_key("chain_damage"))
                {
                    if(attribute["chain_damage"].is_number())
                    {
                        aItem.MagicDamage = attribute["chain_damage"].as_i64().unwrap() as i32;
                        isUselessItem = false;
                    }
                    else if(attribute["chain_damage"].is_string())
                    {
                        aItem.MagicDamage = attribute["chain_damage"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                        isUselessItem = false;
                    }
                }

                // Magic Damage (& bonus_chance_damage) since their calculation is the same
                else if (attribute.contains_key("bonus_chance_damage"))
                {
                    if(attribute["bonus_chance_damage"].is_number())
                    {
                        aItem.MagicDamage = attribute["bonus_chance_damage"].as_i64().unwrap() as i32;
                        isUselessItem = false;
                    }
                    else if(attribute["bonus_chance_damage"].is_string())
                    {
                        aItem.MagicDamage = attribute["bonus_chance_damage"].as_str().unwrap().to_string().parse::<i32>().unwrap();
                        isUselessItem = false;
                    }
                }

                // Magic %
                else if (attribute.contains_key("chain_chance"))
                {
                    if(attribute["chain_chance"].is_number())
                    {
                        aItem.MagicChance_Melee = attribute["chain_chance"].as_i64().unwrap() as f32;
                        isUselessItem = false;
                    }
                    else if(attribute["chain_chance"].is_string())
                    {
                        aItem.MagicChance_Melee = attribute["chain_chance"].as_str().unwrap().to_string().parse::<f32>().unwrap();
                        isUselessItem = false;
                    }
                    aItem.MagicChance_Melee = aItem.MagicChance_Melee / 100.0;
                    aItem.MagicChance_Ranged = aItem.MagicChance_Melee;
                }

                // Magic % (&bonus_chance)
                else if (attribute.contains_key("bonus_chance"))
                {
                    if(attribute["bonus_chance"].is_number())
                    {
                        aItem.MagicChance_Melee = attribute["bonus_chance"].as_i64().unwrap() as f32;
                        isUselessItem = false;
                    }
                    else if(attribute["bonus_chance"].is_string())
                    {
                        aItem.MagicChance_Melee = attribute["bonus_chance"].as_str().unwrap().to_string().parse::<f32>().unwrap();
                        isUselessItem = false;
                    }
                    aItem.MagicChance_Melee = aItem.MagicChance_Melee / 100.0;
                    aItem.MagicChance_Ranged = aItem.MagicChance_Melee;
                }

                // Magic % Bash Melee
                else if (attribute.contains_key("bash_chance_melee"))
                {
                    if(attribute["bash_chance_melee"].is_number())
                    {
                        aItem.MagicChance_Melee = attribute["bash_chance_melee"].as_f64().unwrap() as f32;
                        isUselessItem = false;
                    }
                    else if(attribute["bash_chance_melee"].is_string())
                    {
                        aItem.MagicChance_Melee = attribute["bash_chance_melee"].as_str().unwrap().to_string().parse::<f32>().unwrap();
                        isUselessItem = false;
                    }
                    aItem.MagicChance_Melee = aItem.MagicChance_Melee / 100.0;
                }

                // Magic % Bash Ranged
                else if (attribute.contains_key("bash_chance_ranged"))
                {
                    if(attribute["bash_chance_ranged"].is_number())
                    {
                        aItem.MagicChance_Ranged = attribute["bash_chance_ranged"].as_f64().unwrap() as f32;
                        isUselessItem = false;
                    }
                    else if(attribute["bash_chance_ranged"].is_string())
                    {
                        aItem.MagicChance_Ranged = attribute["bash_chance_ranged"].as_str().unwrap().to_string().parse::<f32>().unwrap();
                        isUselessItem = false;
                    }
                    aItem.MagicChance_Ranged = aItem.MagicChance_Ranged / 100.0;
                }

                // Crit Multiplier
                else if (attribute.contains_key("crit_multiplier"))
                {
                    // bloodthorn crit is activate
                    if(item.0 != "item_bloodthorn")
                    {
                        if(attribute["crit_multiplier"].is_number())
                        {
                            aItem.CritMultiplier = attribute["crit_multiplier"].as_i64().unwrap() as f32;
                            isUselessItem = false;
                        }
                        else if(attribute["crit_multiplier"].is_string())
                        {
                            aItem.CritMultiplier = attribute["crit_multiplier"].as_str().unwrap().to_string().parse::<f32>().unwrap();
                            isUselessItem = false;
                        }
                        aItem.CritMultiplier = aItem.CritMultiplier / 100.0;
                    }
                }

                // Crit %
                else if (attribute.contains_key("crit_chance"))
                {
                    // bloodthorn crit is activate
                    if(item.0 != "item_bloodthorn")
                    {
                        if(attribute["crit_chance"].is_number())
                        {
                            aItem.CritChance = attribute["crit_chance"].as_f64().unwrap() as f32;
                            isUselessItem = false;
                        }
                        else if(attribute["crit_chance"].is_string())
                        {
                            aItem.CritChance = attribute["crit_chance"].as_str().unwrap().to_string().parse::<f32>().unwrap();
                            isUselessItem = false;
                        }
                        aItem.CritChance = aItem.CritChance / 100.0;
                    }
                }
            }
        }

        // Set IsUselessItem
        aItem.IsUselessItem = isUselessItem;

        // If ignore useless items
        if(IgnoreUselessItems && isUselessItem)
        {
            continue;
        }

        // Add to list
        InOutItems.push(aItem);
    }
    
}

fn GetUrlForItem(ItemName: &String) -> String
{
    let mut commonUrl: String = String::from("https://dota2.fandom.com/wiki/");
    let mut specialUrl = ItemName.clone().replace(' ', "_");
    let newUrl = commonUrl + &specialUrl;

    newUrl
}

fn GetItemNames(ItemList: &mut Vec<ItemEntry>)
{
    // Get html source
    let response = reqwest::blocking::get("https://dota2.fandom.com/wiki/Items").unwrap();
    let htmlBody = response.text().unwrap();
    let text = &htmlBody[..];

    let fragment = Html::parse_fragment(text);
    let selectorItemList = Selector::parse("div.itemlist").unwrap();
    let selectorDiv = Selector::parse("div").unwrap();
    let selectorA = Selector::parse("a").unwrap();

    // Find all items by looping all class="itemlist"
    let mut categoryCounter = 0;
    for itemlistElement in fragment.select(&selectorItemList) {
        let mut category = "-1337";
        match categoryCounter {
            0 | 1 | 2 | 3 | 4 => 
            category = "Basics Items",
            5 | 6 | 7 | 8 | 9 | 10 => 
            category = "Upgraded Items",
            11 | 12 | 13 | 14 | 15 | 16 | 17 => 
            category = "Neutral Items",
            18 => 
            category = "Roshan Drop",
            19 => 
            category = "Unreleased Items",
            20 | 21 => 
            category = "Removed Items",
            22 | 23 | 24 | 25 | 26 | 27 | 28 => 
            category = "Event Items",
            _ => category = "No idea",
        }

        for element in itemlistElement.select(&selectorDiv) {
            let mut elementList = element.select(&selectorA);
            // Second a
            let mut e = elementList.next();
            e = elementList.next();

            let itemName = e.unwrap().inner_html();
            let itemEntry = ItemEntry {Name: itemName, Category: category.to_string()};
            ItemList.push(itemEntry)
        }

        categoryCounter = categoryCounter + 1;
    }
}

fn WriteItemsToXlsx(wb: &mut Workbook, InItems: &Vec<Item>)
{
    // Create Sheet
    let mut sheet = wb.create_sheet("Items");

    // Write to Sheet
    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        
        for item in InItems
        {
            WriteItem(sw, &item);
        }
        
        Ok(())
    }).unwrap();
}

fn WriteItem(sw: &mut SheetWriter, InItem: &Item)
{
    sw.append_row(row![
        InItem.Name.clone(),
        InItem.Cost.to_string(),
        InItem.Damage.to_string(),
        InItem.Damage_Melee.to_string(),
        InItem.Damage_Ranged.to_string(),
        InItem.AttackSpeed.to_string(),
        InItem.Str.to_string(),
        InItem.Agi.to_string(),
        InItem.Int.to_string(),
        InItem.ArmorCorruption.to_string(),
        InItem.MagicDamage.to_string(),
        InItem.MagicChance_Melee.to_string(),
        InItem.MagicChance_Ranged.to_string(),
        InItem.CritMultiplier.to_string(),
        InItem.CritChance.to_string(),
        InItem.IsNeutralItem.to_string()
    ]);
}

fn WriteHeroesToXlsx(wb: &mut Workbook, InHeroes: &Vec<Hero>)
{
    // Create Sheet
    let mut sheet = wb.create_sheet("Heroes");

    // Write to Sheet
    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        
        for hero in InHeroes
        {
            sw.append_row(row![
                hero.Name.clone(),
                hero.PrimaryAttribute.to_string(),
                hero.AttackType.to_string(),
                hero.BAT.to_string(),
                hero.BaseAttackSpeed.to_string()
            ]);
        }

        Ok(())
    }).unwrap();
}

/*
fn GetItemStats(JsonData: &String, InOutItem: &mut Item) 
{
    let itemURL = GetUrlForItem(&InOutItem.Name);

    // Get html source
    let response = reqwest::blocking::get(itemURL).unwrap();
    let htmlBody = response.text().unwrap();
    let text = &htmlBody[..];

    let fragment = Html::parse_fragment(text);
    let selectorInfoBox = Selector::parse("table.infobox").unwrap();
    let selectorTr = Selector::parse("tr").unwrap();
    let selectorDivDiv = Selector::parse("div").unwrap();

    // Find stats (infobox)
    let InfoBox = fragment.select(&selectorInfoBox).next().unwrap();

    // Find cost
    let mut costTrIt = InfoBox.select(&selectorTr);
    
    // Get fourth tr
    let mut costTr = costTrIt.next();
    costTr = costTrIt.next();
    costTr = costTrIt.next();
    costTr = costTrIt.next();

    let mut divdivIt = costTr.unwrap().select(&selectorDivDiv);

    // Second div is the cost
    let mut costDiv = divdivIt.next();
    costDiv = divdivIt.next();
    let costString = costDiv.unwrap().inner_html();

    

    // Get the cost as int
    let searchStr = "Cost<br>";
    let searchStrLen = searchStr.chars().count();
    let costIndex = costString.find(searchStr).unwrap();

    // Find cost xxxx end
    let searchStr2 = " ";
    let costEndIndex = costString.find(searchStr2).unwrap();

    let str = &costString[costIndex+searchStrLen..costEndIndex].to_string();
    let costAsInt = str.parse::<i32>().unwrap();

    // Find stats (bonus)
    let selectorInt = Selector::parse("div").unwrap();

    //let bonus = 


    println!("{:#?}", costAsInt);
}*/