#!/usr/bin/env python

import pandas as pd
import os
import time
import re
import glob

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from openpyxl import load_workbook


"""
HLL Stats Scraper

The Script goes through several private and public RCONS from game Hell Let Loose
and retrieves data from best players performance in several matches aggrouping it 
in a .xlsx file in the end.

To debug code remove headless line to see normal script run.
"""

####################################################################

# ARGUMENTS
#KILLS_PER_MINUTE_REQ = 0.8
KILLS_REQ = 65
GAME_ID_SEARCH_NUMBER = 8500

####################################################################




HLL_PRIVATE_RCONS = {
  #"EXD": "http://164.90.235.39:8010/#/gamescoreboard/", rcon changed
  #"EXD_2": "http://164.90.235.39:8012/#/gamescoreboard/", rcon changed
  #"EXD_3": "http://164.90.235.39:8011/#/gamescoreboard/", rcon changed
  "WTH": "http://144.91.71.61:8040/#/gamescoreboard/", #WTH_1
  #"WTH_2": "http://144.91.71.61:8041/#/gamescoreboard/",
  "Polish_Crew_1": "http://178.18.254.195:8010/#/gamescoreboard/",
  #"Polish_Crew_2": "http://85.190.254.165:8010/#/gamescoreboard/",
  "PHX": "http://teamphx.dynv6.net:8010/#/gamescoreboard/", #PHX_1
  #"PHX_2": "http://rcon.team-phx.com:8011/#/gamescoreboard/",
  "ESPT": "http://144.126.151.248:8010/#/gamescoreboard/" #ESPT_1
}

HLL_PUBLIC_RCONS = {
  #"82AD_2_test": "https://server2.82nd.gg/#/gamescoreboard/25982",
  #"82AD_1": "https://server1.82nd.gg/#/gamescoreboard/"
  
  #"82AD_3": "https://server3.82nd.gg/#/gamescoreboard/", # event server
  #"HLL_PL_1": "https://stats1.hellletloose.pl/#/gamescoreboard/",
  "82AD": "https://server2.82nd.gg/#/gamescoreboard/", #82AD_2
  #"HLL_PL_2": "https://stats2.hellletloose.pl/#/gamescoreboard/",
  "ESPT_2": "http://144.126.151.248:7014/#/gamescoreboard/",
  #"Saucymuffin": "http://50.116.61.211:7010/#/gamescoreboard/" rcon changed
  "Glow": "http://www.marech.fr:7050/#/gamescoreboard/"
}

####################################################################

def clean_csv(df):
    df['Name'] = df['Name'].str.replace('Name', '', 1)
    df['Kills'] = df['Kills'].str.replace('Kills', '', 1)
    df['Deaths'] = df['Deaths'].str.replace('Deaths', '', 1)
    df['K/D'] = df['K/D'].str.replace('K/D', '', 1)
    df['Max kill streak'] = df['Max kill streak'].str.replace('Max kill streak', '', 1)
    df["Kill(s) / minute"] = df["Kill(s) / minute"].str.replace("Kill\(s\) / minute", '', 1, regex=True)
    df['Death(s) / minute'] = df['Death(s) / minute'].str.replace('Death\(s\) / minute', '', regex=True)
    df['Max death streak'] = df['Max death streak'].str.replace('Max death streak', '', 1)
    df['Max TK streak'] = df['Max TK streak'].str.replace('Max TK streak', '', 1)
    df['Death by TK'] = df['Death by TK'].str.replace('Death by TK', '', 1)
    df['Death by TK Streak'] = df['Death by TK Streak'].str.replace('Death by TK Streak', '', 1)
    df['(aprox.) Longest life min.'] = df['(aprox.) Longest life min.'].str.replace('\(aprox.\) Longest life min.', '', 1, regex=True)
    df['(aprox.) Shortest life secs.'] = df['(aprox.) Shortest life secs.'].str.replace('\(aprox.\) Shortest life secs.', '', 1, regex=True)
    df['Nemesis'] = df['Nemesis'].str.replace('Nemesis', '', 1)
    df['Victim'] = df['Victim'].str.replace('Victim', '', 1)
    df['Weapons'] = df['Weapons'].str.replace('Weapons', '', 1)
    return df

def parse_table_data():
    df = pd.read_html(driver.page_source)
    df = df[0]
    df = clean_csv(df)

    df["game_id"] = driver.current_url

    df = df.astype({'Kill(s) / minute':'float'})
    df = df.astype({'Kills':'int'})
    #df = df.loc[(df['Kill(s) / minute'] >= KILLS_PER_MINUTE_REQ) & (df['Kills'] >= KILLS_REQ)]
    df = df.loc[df['Kills'] >= KILLS_REQ]

    return df

def get_table_data_from_priv_server():
    kill_sort = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div[1]/div/div[2]/div/div[3]/table/thead/tr/th[2]/span")
    kill_sort.click()
    kill_sort.click()

def click_more_stats():
    more_stats = driver.find_element(By.XPATH, "/html/body/div/div/div[3]/div[17]/div/div/button")
    more_stats.click()

def get_table_data_from_pub_server():
    #kill_sort = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[3]/div[17]/div/div[2]/div/div[3]/table/thead/tr/th[2]")
    try:
        kill_sort = WebDriverWait(driver,3).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div/div[3]/div[17]/div/div[2]/div/div[3]/table/thead/tr/th[2]")))
        kill_sort.click() 
        kill_sort.click()
    except:
        pass

def remove_duplicate_values(df):
    df2 = df.drop_duplicates(subset=["Name", "Kills", "K/D", "Weapons"], keep='first')
    return df2

def create_excel():
    first_entry = True
    for file in glob.glob("*.csv"):
        df = pd.read_csv(file)

        df = remove_duplicate_values(df)

        name = file.split('.')
        filename = name[0].split('/')
        filename = filename[-1]

        #create xlsx
        if first_entry:
            with pd.ExcelWriter('hll_stats.xlsx', mode='w') as writer:  
                df.to_excel(writer, sheet_name=filename, index=False)

        #append    
        else:
            excel_path = "hll_stats.xlsx"
            ExcelWorkbook = load_workbook(excel_path)
            writer = pd.ExcelWriter(excel_path, engine = 'openpyxl')
            writer.book = ExcelWorkbook
            df.to_excel(writer, sheet_name = filename, index=False)
            writer.save()
            writer.close()

        first_entry = False

def clean_csvs_and_excel():
    files=glob.glob('*.csv')
    for filename in files:
        os.unlink(filename)

    files=glob.glob('*.xlsx')
    for filename in files:
        os.unlink(filename)

def load_private_rcons():
    for server_name in HLL_PRIVATE_RCONS:
        driver.get(HLL_PRIVATE_RCONS[server_name])
        time.sleep(1)
        driver.maximize_window()

        first_entry_header = True

        if first_entry_header:
            time.sleep(3)

        current_game_id = driver.current_url.rpartition('/')[2]
        game_url_default = re.sub(r'\d+$', '', driver.current_url)
        

        

        try:
            current_game_id = int(current_game_id)
            click_more_stats()
        except:
            print("No stats for this link: ", server_name + " - " + driver.current_url)

        for i in range(GAME_ID_SEARCH_NUMBER):
            try:
                print("Private RCON: " + server_name + " - " + game_url_default + str(current_game_id-i))
                if (current_game_id-i) <= 0:
                    break
                
                if (i != 0):
                    driver.get(game_url_default + str(current_game_id-i))

                time.sleep(1)
                driver.maximize_window()

                get_table_data_from_priv_server()
                df = parse_table_data()

                if not df.empty:
                    if first_entry_header:
                        df.to_csv(server_name + '.csv', mode='a', header=True, index=False)
                    else:
                        df.to_csv(server_name + '.csv', mode='a', header=False, index=False)
                    first_entry_header = False

            except:
                print("No stats for this link: ", server_name + " - " + game_url_default + str(current_game_id-i))

def load_public_rcons():
    for server_name in HLL_PUBLIC_RCONS:
        driver.get(HLL_PUBLIC_RCONS[server_name])
        time.sleep(2)
        driver.maximize_window()

        first_entry_header = True

        if first_entry_header:
            time.sleep(3)
        

        current_game_id = driver.current_url.rpartition('/')[2]
        game_url_default = re.sub(r'\d+$', '', driver.current_url)
        
        

        try:
            current_game_id = int(current_game_id)
            click_more_stats()
        except:
            print("No stats for this link: ", server_name + " - " + driver.current_url)

        
        for i in range(GAME_ID_SEARCH_NUMBER):
            print("Public RCON: " + server_name + " - " + game_url_default + str(current_game_id-i))
            try:
                if (current_game_id-i) <= 0:
                    break

                if (i != 0):
                    driver.get(game_url_default + str(int(current_game_id-i)))

                time.sleep(1)
                #driver.maximize_window()
                get_table_data_from_pub_server()
                df = parse_table_data()

                if not df.empty:
                    if first_entry_header:
                        df.to_csv(server_name + '.csv', mode='a', header=True, index=False)
                    else:
                        df.to_csv(server_name + '.csv', mode='a', header=False, index=False)
                    first_entry_header = False

            except ValueError as err:
                print("No stats for this link: ", server_name + " - " + game_url_default + str(current_game_id-i))
                print(err)

def exit_functions():
    driver.close()
    driver.quit()

def init_webdriver():
    options = FirefoxOptions()
    options.add_argument("--headless")
    driver = webdriver.Firefox(options=options)
    return driver

if __name__ == "__main__":
    clean_csvs_and_excel()

    driver = init_webdriver()

    load_private_rcons()
    load_public_rcons()

    exit_functions()
    
    create_excel()
