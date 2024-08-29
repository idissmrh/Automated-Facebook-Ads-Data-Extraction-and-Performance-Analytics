import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import numpy as np
import os

# Function to export Facebook Ads data using Selenium

start_date = '2024-08-24'
end_date = '2024-08-29'
start_date_file = 'Aug-24-2024'
end_date_file = 'Aug-28-2024'


def export_facebook_ads_data(username, password, link, file_path, start_date, end_date, start_date_file, end_date_file):
    time_range = f"{start_date}_{end_date}"
    time_range_files = f"{start_date_file}-to-{end_date_file}"
    link = link.replace("2024-08-01_2024-08-24", time_range)
    file_path = file_path.replace(
        'Aug-5-2024-to-Aug-23-2024', time_range_files)
    try:
        # Use webdriver_manager to automatically download and set up ChromeDriver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
    except Exception as e:
        print(f"Automatic setup failed: {e}")
        # If automatic setup fails, use a manually downloaded ChromeDriver
        chrome_driver_path = '/path/to/chromedriver'
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service)

    # Open Facebook Ads Manager
    driver.get(link)

    # Log in to Facebook
    email_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'email')))
    email_input.send_keys(username)
    password_input = driver.find_element(By.ID, 'pass')
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)

    # Wait for the Ads Manager page to load
    # Adjust this sleep time as necessary for my internet speed
    time.sleep(30)

    # Click the "Export table data" button
    try:
        export_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[1]/div/div/div/div[1]/div/span/div/div[1]/div[3]/div[2]/span/div/div/div/div/div[1]/div/div/div[2]/div/span/div/div')))
        export_button.click()
        print("Export button clicked")
    except Exception as e:
        print(f"Failed to click export button: {e}")
        driver.quit()
        exit()

    # Wait for the export format options to appear and click the CSV option
    time.sleep(30)
    try:
        export_button = driver.find_element(
            By.XPATH, "//div[contains(@class, 'x1i10hfl xjqpnuy xa49m3k xqeqjp1 x2hbi6w x972fbf xcfux6l x1qhh985 xm0m39n x9f619 x1ypdohk xe8uvvx xdj266r x11i5rnm xat24cr x1mh8g0r x2lwn1j xeuugli x16tdsg8 xggy1nq x1ja2u2z x1t137rt x6s0dn4 x1ejq31n xd10rxx x1sy0etr x17r0tee x3nfvp2 xdl72j9 x1q0g3np x2lah0s x193iq5w x1n2onr6 x1hl2dhg x87ps6o xxymvpz xlh3980 xvmahel x1lku1pv xhk9q7s x1otrzb0 x1i1ezom x1o6z2jb x1xlr1w8 x140t73q xb57al4 x1y1aw1k xwib8y2 x1swvt13 x1pi30zi')]//div[contains(text(), 'Export')]")

        export_button.click()
        print("Export button 2 clicked")
    except Exception as e:
        print(f"Failed to click the download button: {e}")
        driver.quit()
        exit()

    # Wait for the download to complete (adjust sleep time as necessary)
    time.sleep(50)

    # Close the browser
    driver.quit()
    print("Data exported successfully.")
    return file_path


# Function to standardize and analyze the exported data


def Standarize_analysis(path, col):
    df = pd.read_excel(path)

    df = df[df['Campaign name'].str.contains('CBB', na=False)]
    df = df.fillna(0)
    df = df.groupby(['Account name', 'Reporting starts', 'Reporting ends', 'Ad name', 'Campaign name', 'Ad Set Name', col]).agg({

        'Impressions': 'sum',
        'Amount spent (USD)': 'sum',
        'Link clicks': 'sum',
        'Results': 'sum'
    })
    df = df.rename(columns={'Results': 'Appointments booked'})
    df['Filter'] = col

    df.replace(float('inf'), float(0.0), inplace=True)

    return df


def Standarize_analysis2(path):
    df = pd.read_excel(path)

    df = df[df['Campaign name'].str.contains('CBB', na=False)]
    df = df.fillna(0)
    df = df.groupby(['Account name', 'Reporting starts', 'Reporting ends', 'Ad name', 'Campaign name', 'Ad Set Name']).agg({


        'Impressions': 'sum',
        'Amount spent (USD)': 'sum',
        'Link clicks': 'sum',
        'Results': 'sum'
    })
    df = df.rename(columns={'Results': 'Appointments booked'})

    df.replace(float('inf'), float(0.0), inplace=True)

    return df


def Standarize_analysis3(path):
    df = pd.read_excel(path)

    df = df[df['Campaign name'].str.contains('CBB', na=False)]
    df = df.fillna(0)
    df = df.groupby(['Ad name', 'Campaign name', 'Ad Set Name']).agg({

        'Video plays at 25%': 'sum',
        'Video plays at 50%': 'sum',
        'Video plays at 75%': 'sum',
        'Video plays at 95%': 'sum',
        'Video plays at 100%': 'sum',

        'Impressions': 'sum',
        'Amount spent (USD)': 'sum',
        'Link clicks': 'sum',
        'Results': 'sum'
    })
    df = df.rename(columns={'Results': 'Appointments booked'})
    df['25% to 50% Video Length'] = np.abs(
        (df['Video plays at 25%']-df['Video plays at 50%'])/(df['Video plays at 25%']))
    df['50% to 75% Video Length'] = np.abs(
        (df['Video plays at 50%']-df['Video plays at 75%'])/(df['Video plays at 50%']))
    df['75% to 95% Video Length'] = np.abs(
        (df['Video plays at 75%']-df['Video plays at 95%'])/(df['Video plays at 75%']))
    df['95% to 100% Video Length'] = np.abs(
        (df['Video plays at 95%']-df['Video plays at 100%'])/(df['Video plays at 95%']))
    df['25% Video Length'] = np.abs(
        (df['Video plays at 25%'])/(df['Impressions']))
    df['50% Video Length'] = np.abs(
        (df['Video plays at 50%'])/(df['Impressions']))
    df['75% Video Length'] = np.abs(
        (df['Video plays at 75%'])/(df['Impressions']))
    df['95%  Video Length'] = np.abs(
        (df['Video plays at 95%'])/(df['Impressions']))
    df['100%  Video Length'] = np.abs(
        (df['Video plays at 100%'])/(df['Impressions']))

    df.replace(float('inf'), float(0.0), inplace=True)
    df.reset_index(df.index.names, inplace=True)
    df.drop(['Video plays at 25%',
             'Video plays at 50%', 'Video plays at 75%', 'Video plays at 95%',
             'Video plays at 100%'], axis=1, inplace=True)

    # Melt the DataFrame for columns_to_melt_1
    melted_1 = df.melt(id_vars=['Ad name', 'Campaign name', 'Ad Set Name'],
                       value_vars=columns_to_melt_1, var_name='Drop-off Rate Stage', value_name='Drop-off Rate')

# Melt the DataFrame for columns_to_melt_2
    melted_2 = df.melt(id_vars=['Ad name', 'Campaign name', 'Ad Set Name'],
                       value_vars=columns_to_melt_2, var_name='Video Play Stage', value_name='Video Plays')
    melted_2.rename(columns={'Ad name': 'Ad name_2', 'Campaign name': 'Campaign name_2',
                    'Ad Set Name': 'Ad Set Name 2'}, inplace=True)

# Combine the melted DataFrames
    combined_df = pd.concat([melted_1, melted_2], axis=1)
    combined_df.drop(['Campaign name_2', 'Ad Set Name 2',
                     'Ad name_2'], axis=1, inplace=True)
    combined_df['Drop-off Rate'] = combined_df['Drop-off Rate'].fillna(0)
    return combined_df


# Function to save the analysis results to a CSV file

# Main script to integrate everything
username = '**********@gmail.com'
password = '*****'
file_path_1 = 'C:/Users/mirah/Downloads/Untitled-report-Aug-5-2024-to-Aug-23-2024.xlsx'
file_path_2 = 'C:/Users/mirah/Downloads/Untitled-report-Aug-5-2024-to-Aug-23-2024 (1).xlsx'
file_path_3 = 'C:/Users/mirah/Downloads/Untitled-report-Aug-5-2024-to-Aug-23-2024 (2).xlsx'
file_path_4 = 'C:/Users/mirah/Downloads/Untitled-report-Aug-5-2024-to-Aug-23-2024 (3).xlsx'
link_1 = 'https://adsmanager.facebook.com/adsmanager/reporting/view?act=327267577&business_id=172660837235742&event_source=CLICK_CREATE_REPORT&breakdowns=ad_name%2Ccampaign_name%2Cadset_name&empty_comparison_time_range=true&filter_set=had_delivery-STRING%1EEQUAL%1E%221%22&locked_dimensions=2&metrics=impressions%2Cattribution_setting%2Cresults%2Cspend%2Ccost_per_result%2Cdelivery_start%2Cactions%3Alink_click%2Cdate_start%2Cdate_stop%2Caccount_name&time_range=2024-08-01_2024-08-24%2C&target_currency=USD'
link_2 = 'https://adsmanager.facebook.com/adsmanager/reporting/view?act=327267577&business_id=172660837235742&event_source=CLICK_CREATE_REPORT&breakdowns=ad_name%2Ccampaign_name%2Cadset_name%2Cage%2Cgender&empty_comparison_time_range=true&filter_set=had_delivery-STRING%1EEQUAL%1E%221%22&locked_dimensions=4&metrics=impressions%2Cattribution_setting%2Cresults%2Cspend%2Ccost_per_result%2Cdelivery_start%2Cactions%3Alink_click%2Cdate_start%2Cdate_stop%2Caccount_name&time_range=2024-08-01_2024-08-24%2C&target_currency=USD'
link_3 = 'https://adsmanager.facebook.com/adsmanager/reporting/view?act=327267577&business_id=172660837235742&event_source=CLICK_CREATE_REPORT&breakdowns=ad_name%2Ccampaign_name%2Cadset_name%2Cpublisher_platform%2Cplatform_position%2Cdevice_platform&empty_comparison_time_range=true&filter_set=had_delivery-STRING%1EEQUAL%1E%221%22&locked_dimensions=5&metrics=impressions%2Cattribution_setting%2Cresults%2Cspend%2Ccost_per_result%2Cdelivery_start%2Cactions%3Alink_click%2Cdate_start%2Cdate_stop%2Caccount_name&time_range=2024-08-01_2024-08-24%2C&target_currency=USD'
link_4 = 'https://adsmanager.facebook.com/adsmanager/reporting/view?act=327267577&business_id=172660837235742&event_source=CLICK_CREATE_REPORT&breakdowns=ad_name%2Ccampaign_name%2Cadset_name&filter_set=had_delivery-STRING%1EEQUAL%1E%221%22&locked_dimensions=2&metrics=impressions%2Cattribution_setting%2Cresults%2Cspend%2Ccost_per_result%2Cdelivery_start%2Cvideo_p25_watched_actions%3Avideo_view%2Cvideo_p50_watched_actions%3Avideo_view%2Cvideo_p75_watched_actions%3Avideo_view%2Cvideo_p95_watched_actions%3Avideo_view%2Cvideo_p100_watched_actions%3Avideo_view%2Cactions%3Alink_click&time_range=2024-08-01_2024-08-24%2C&target_currency=USD&breakdown_regrouping=1'
columns_to_melt_1 = [
    '25% to 50% Video Length',
    '50% to 75% Video Length',
    '75% to 95% Video Length',
    '95% to 100% Video Length'

]
columns_to_melt_2 = [
    '25% Video Length',
    '50% Video Length',
    '75% Video Length',
    '95%  Video Length',
    '100%  Video Length'
]

exported_file_path_1 = export_facebook_ads_data(
    username, password, link_1, file_path_1, start_date, end_date, start_date_file, end_date_file)

exported_file_path_2 = export_facebook_ads_data(
    username, password, link_2, file_path_2, start_date, end_date, start_date_file, end_date_file)

exported_file_path_3 = export_facebook_ads_data(
    username, password, link_3, file_path_3, start_date, end_date, start_date_file, end_date_file)

exported_file_path_4 = export_facebook_ads_data(
    username, password, link_4, file_path_4, start_date, end_date, start_date_file, end_date_file)
# Analyze the exported data
df_Ad_name = pd.DataFrame(Standarize_analysis2(exported_file_path_1))
df_Ad_name.reset_index(df_Ad_name.index.names, inplace=True)
df_platform = pd.DataFrame(Standarize_analysis(
    exported_file_path_3, 'Platform'))
df_placement = pd.DataFrame(Standarize_analysis(
    exported_file_path_3, 'Placement'))
df_dev_platform = pd.DataFrame(Standarize_analysis(
    exported_file_path_3, 'Device platform'))

df_Age = pd.DataFrame(Standarize_analysis(exported_file_path_2, 'Age'))
df_Gender = pd.DataFrame(
    Standarize_analysis(exported_file_path_2, 'Gender'))

df = pd.concat([df_Age, df_Gender, df_platform,
               df_placement, df_dev_platform])
df.reset_index(df.index.names, inplace=True)
df_video = pd.DataFrame(Standarize_analysis3(exported_file_path_4))


# Save the analysis results
file_name = 'C:/Users/mirah/Downloads/final1.xlsx'

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    df_Ad_name.to_excel(writer, sheet_name='Sheet1', index=False)
    df.to_excel(writer, sheet_name='Sheet2', index=False)
    df_video.to_excel(writer, sheet_name='Sheet3', index=False)

print("Data analysis and export completed successfully.")

time.sleep(120)

# Remove the file after 5 minutes
file_paths = [exported_file_path_1, exported_file_path_2,
              exported_file_path_3, exported_file_path_4]
for file_path in file_paths:
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"{file_path} has been removed after 2 minutes.")
    else:
        print(f"{file_path} not found, could not remove it.")

exit()
