from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from dotenv import load_dotenv
import os
from time import sleep

load_dotenv()

usa_portal_url = os.environ.get('ANET_US_PORTAL_URL')
can_portal_url = os.environ.get('ANET_CAN_PORTAL_URL')

download_dir = os.path.join(os.getcwd(), 'files')
downloaded_file_count = 0


def prepare_chrome_browser():
    # Prepare and return an instance of a chrome browser object
    chrome_options = webdriver.ChromeOptions()

    # Add Download configuration
    prefs = {'download.default_directory' : download_dir}
    chrome_options.add_experimental_option('prefs', prefs)

    # Use Browser in Headless mode
    chrome_options.headless = True
    browser = webdriver.Chrome(options=chrome_options)
    return browser


def login_anet_portal(browser, portal):
    # Login to the ActiveNet portal with creds from the .env file
    username = os.environ.get('USERNAME', None)
    password = os.environ.get('PASSWORD', None)
    if portal == 'USA':
        url = usa_portal_url
    elif portal == 'CAN':
        url = can_portal_url
    else:
        raise Exception(f"Unable to determine portal to log into from {portal}")

    browser.get(url)
    browser.find_element(By.ID, 'oLogin_UserName').send_keys(username)
    browser.find_element(By.ID, 'oLogin_Password').send_keys(password)
    browser.find_element(By.ID, 'oLogin_LoginButton').click()


def search_organization(browser, org_name):
    # Open the Organization search dialog and search for the provided org name
    print(f'Searching for {org_name}')
    window_before = browser.window_handles[0]
    browser.find_element(By.ID, 'ctl00_cntPlhd_btnCustomSearch').click()
    window_after = browser.window_handles[1]
    browser.switch_to.window(window_after)
    browser.find_element(By.ID, 'ctl00_cntPlhd_txtOrgURL').send_keys(org_name)
    browser.find_element(By.ID, 'ctl00_cntPlhd_btnSearch').click()
    browser.switch_to.window(window_before)


def select_organization(browser, org_name, production=True):
    # Find the organization in the list and check its do stuff checkbox
    print('Selecting the Do Stuff checkbox for the found org')
    url_link = browser.find_element(By.LINK_TEXT, org_name)
    url_link_id = url_link.get_attribute('id')
    url_link_parts = url_link_id.split('_')
    url_link_parts[2] = 'DtDoStf'
    url_link_parts[4] = 'chkOrgID'
    checkbox_id = '_'.join(url_link_parts[:5])
    browser.find_element(By.ID, checkbox_id).click()
    browser.find_element(By.ID, 'ctl00_cntPlhd_btnDoStuff').click()


def wait_for_download_finish():
    global downloaded_file_count
    wait = True
    current_file_count = downloaded_file_count
    while(wait and current_file_count == downloaded_file_count):
        print('waiting for download to finish')
        sleep(2)
        for fname in os.listdir(download_dir):
            if 'active_report' in fname and not 'crdownload' in fname:
                wait = False
                print("Downloaded Files: ", downloaded_file_count)
                downloaded_file_count = downloaded_file_count + 1
                break


def run_query(browser, name, query, production=True):
    global downloaded_file_count
    # Select either to run on production or trainer
    if production:
        browser.find_element(By.ID, 'ctl00_cntPlhd_rdSiteSelection_1').click()
    else:
        browser.find_element(By.ID, 'ctl00_cntPlhd_rdSiteSelection_2').click()

    # Select to output the results to a csv file
    Select(browser.find_element(By.ID, 'ctl00_cntPlhd_OutputType')).select_by_value("2")

    # Drop the query in the SQL Command text box
    browser.find_element(By.ID, 'ctl00_cntPlhd_SQLCommand').clear()
    browser.find_element(By.ID, 'ctl00_cntPlhd_SQLCommand').send_keys(query)
    window_before = browser.window_handles[0]

    # Run the query
    print(f'Running query {name}')
    browser.find_element(By.ID, 'ctl00_cntPlhd_btnGoExecuteSQL').click()

    current_file_count = downloaded_file_count

    # If we opened a new window then handle the new window
    if len(browser.window_handles) > 1:
        window_after = browser.window_handles[1]
        browser.switch_to.window(window_after)
        wait_for_download_finish()
        browser.close()
        browser.switch_to.window(window_before)

    # Otherwise check to see if we get an error or a download on the main page
    else:
        error_box = browser.find_element(By.ID, 'ctl00_lblErrMsg')
        if error_box.text != '':
            return                  # Return here if we get any error message
        wait_for_download_finish()
    
    # Find the downloaded file and rename it to report_name.csv
    sleep(2)
    files = os.listdir(download_dir)
    paths = [os.path.join(download_dir, basename) for basename in files]
    file_path = max(paths, key=os.path.getctime)
    print(file_path)
    if 'active_report' in file_path and downloaded_file_count > current_file_count:
        os.rename(file_path, os.path.join(download_dir, name))


def main():
    org_name = 'LJSupport12' #'Chicagoparkdistrict'
    portal = 'USA'

    browser = prepare_chrome_browser()
    login_anet_portal(browser, portal)    
    search_organization(browser, org_name)
    select_organization(browser, org_name)

    season_id = 5 # 63
    queries = {
        "facility_listing.csv": "SELECT [Facilities].FacilityNumber, [Facilities].FacilityName, [FacilityTypes].Description as FacilityType, [Facilities].[Retired] as Retired_Facilities,[Facilities].[CENTER_ID], [CENTERS].[CENTERNAME], [CENTERS].[SITE_ID],[SITES].[SITENAME],[GEOGRAPHIC_AREAS].[GEOGRAPHIC_AREAS_NAME],[CENTERS].[ADDRESS1] as Centers_Address1,[CENTERS].[Retired] as Retired_Centers FROM [Facilities] left join [CENTERS] on Facilities.Center_id = Centers.Center_id left join [SITES] on [CENTERS].[SITE_ID]=[SITES].[SITE_ID] left join [GEOGRAPHIC_AREAS] ON [SITES].[GEOGRAPHIC_AREA_ID] = [GEOGRAPHIC_AREAS].[GEOGRAPHIC_AREA_ID] left join [FacilityTypes] on FacilityTypes.facilitytype_id = Facilities.facilitytype_id ORDER BY [FacilityName]",
        "residency_and_checklist_requirements.csv": f"SELECT DISTINCT ACTIVITIES.ACTIVITYNAME ,ACTIVITIES.ACTIVITYNUMBER ,SEASONS.SEASONNAME , CASE ACTIVITIES.RESIDENCY_TYPE WHEN 0 THEN 'BOTH' WHEN 1 THEN 'RESIDENT ONLY' WHEN 2 THEN 'NON-RESIDENT ONLY' END AS 'RESIDENCY TYPE' ,STAGES.DESCRIPTION AS'CHECKLIST_NAME' , CASE ATTACHEDCHECKLISTITEMS.ISREQUIRED WHEN 0 THEN 'NO' WHEN -1 THEN 'YES' END AS 'REQUIRED?' ,CASE ATTACHEDCHECKLISTITEMS.SHOWONLINE WHEN 0 THEN 'NO' WHEN -1 THEN 'YES' END AS 'SHOW ONLINE?' , CASE ATTACHEDCHECKLISTITEMS.INCLUDECONDITION WHEN 0 THEN 'ALWAYS' WHEN 1 THEN 'Resident' WHEN 2 THEN 'Non-resident' WHEN 3 THEN 'Minor' WHEN 4 THEN 'Senior' WHEN 5 THEN 'Member' WHEN 6 THEN 'Non-member' WHEN 7 THEN 'Internet Sale' WHEN 8 THEN 'age is less than or equal to' END AS 'INCLUDE?' from ACTIVITIES ACTIVITIES (nolock) join ATTACHEDCHECKLISTITEMS ATTACHEDCHECKLISTITEMS on ATTACHEDCHECKLISTITEMS.ACTIVITY_ID= ACTIVITIES.ACTIVITY_ID join STAGES STAGES on STAGES.STAGE_ID= ATTACHEDCHECKLISTITEMS.STAGE_ID JOIN SEASONS SEASONS ON SEASONS.SEASON_ID=ACTIVITIES.SEASON_ID where ACTIVITIES.SEASON_ID IN ({season_id})",
        "revenue_details_activities.csv": f"select site.Sitename, a.activityname, a.activitynumber, at.description as activitytype, s.seasonname, c.CHARGE_NAME, af.CHARGE_NAME, Case af.CHARGE_TYPE when 0 then 'Fee' when 1 then 'Discount' End as CHARGETYPE, g.ACCOUNTNAME, c.site_id, af.feeamount, ct.description as customertype, af.overrideflag, af.DISCOUNT_GROUP_NUMBER, EXCLUDE_OTHER_DISCOUNTS_IN_GROUP from activities a (nolock) left join sites site on site.site_id=a.site_id left join seasons s on s.SEASON_ID=a.SEASON_ID left join activity_types at on at.activity_type_id=a.activity_type_id left join ACTIVITy_FEES af on af.ACTIVITY_ID = a.ACTIVITY_ID left join GLACCOUNTS g on g.GLACCOUNT_ID = af.GLACCOUNT_ID left join charges c on af.charge_id=c.charge_id left join customertypes ct on ct.customertype_id=af.customertype_id where a.season_id={season_id}",
        "membership_discount_qualifications.csv": f"select a.activityname, a.activitynumber, at.description as activity_type, s.seasonname, ac.categoryname, si.sitename, af.CHARGE_NAME, case af.CHARGE_TYPE when 0 then 'Fee' when 1 then 'Discount' END as 'CHARGE_TYPE', af.FEEORDER, af.FEEAMOUNT, af.DISCOUNTPERCENT, case af.PREFILLCONDITION when 0 then 'Never' when 1 then 'Always' when 2 then 'If Resident' when 3 then 'If Non-resident' when 4 then 'If Minor' when 5 then 'If Senior' when 6 then 'If Internet' when 7 then 'If Member' when 8 then 'If Non-member' END as PREFILCONDITION, p.packagename, af.DISCOUNT_ORDER, af.DISCOUNT_GROUP_NUMBER from activities a (nolock) left join activity_types at on at.ACTIVITY_TYPE_ID=a.ACTIVITY_TYPE_ID left join seasons s on s.season_id=a.season_id left join rg_category ac on ac.rg_category_id=a.rg_category_id left join sites si on si.site_id=a.site_id left join ACTIVITY_FEES af on af.ACTIVITY_ID= a.ACTIVITY_ID left join activityfeepackages afp on afp.ACTIVITY_FEE_ID=af.activity_fees_id left join packages p on p.PACKAGE_ID=afp.PACKAGE_ID where a.season_id in ({season_id}) order by a.activitynumber asc",
        "membership_package_detail.csv": "SELECT [PACKAGE_ID],[PACKAGENAME] ,(select pc.CATEGORYNAME from PACKAGE_CATEGORIES pc where pc.PACKAGECATEGORY_ID=p.PACKAGECATEGORY_ID) category_name ,(select s.SITENAME from SITES s where s.site_ID=p.site_ID) SITE_NAME ,[AGESMIN] ,[AGESMAX] ,[CATALOGDESCRIPTION] ,[DESCRIPTION] ,[MAXPASSES] ,[MAXUSES] ,case p.Packagestatus when 0 then 'Open' when 1 then 'Closed' end as 'PACKAGE_STATUS' ,case p.RENEWABLEMEMBERSHIP when 0 then 'NO' when -1 then 'YES' end as 'RENEWABLEMEMBERSHIP' ,[PACKAGE_START_DATE] ,[SPECIFICENDDATE] ,case p.MEMBERSHIPDISCOUNT when 0 then 'NO' when -1 then 'YES' end as 'MEMBERSHIPDISCOUNT' ,case p.HIDEONINTERNET when 0 then 'NO' when -1 then 'YES' end as 'HIDEONINTERNET' ,case p.AVAILABLE_AS_PREREQUISITE when 0 then 'NO' when -1 then 'YES' end as 'AVAILABLE_AS_PREREQUISITE' ,case p.NO_EXPIRY when 0 then 'NO' when -1 then 'YES' end as 'NO_EXPIRY' ,[MAX_TIMEPERIODS] ,case p.AUTO_ALLOCATE_TO_IMMEDIATE_FAMILY_MEMBERS when 0 then 'NO' when -1 then 'YES' end as 'AUTO_ALLOCATE_TO_IMMEDIATE_FAMILY_MEMBERS' ,case p.QUALIFY_MEMBER_REGISTRATION_DATES when 0 then 'NO' when -1 then 'YES' end as 'QUALIFY_MEMBER_REGISTRATION_DATES' FROM [PACKAGES] p (nolock)",
        "activity_package_detail.csv": f"SELECT S.SITENAME, a1.ACTIVITYNUMBER, a1.ACTIVITYNAME 'Pkg Name', a1.ENROLLMIN 'Pkg Target', a1.ENROLLMAX 'Pkg Max', A1.MAXENROLLEDONLINE 'Pkg max online', CASE(a1.NOINTERNETREG) WHEN '-1' THEN 'YES' WHEN '0' THEN 'NO' WHEN '1' THEN 'YES' END 'Pkg No internet', ap.DESCRIPTION 'Activity list name', CASE (ap.ENROLL_ALL_ACTIVITIES) WHEN '-1' THEN 'YES' WHEN '0' THEN 'NO' WHEN '1' THEN 'YES' END 'Enroll All?', a2.ACTIVITYNUMBER 'List ActNum', a2.ACTIVITYNAME 'Activity List Activities', a2.SEASON_ID 'List Season', a2.ENROLLMIN 'List Target', a2.ENROLLMAX 'List Max', A2.MAXENROLLEDONLINE 'List max online', CASE(a2.NOINTERNETREG) WHEN '-1' THEN 'YES' WHEN '0' THEN 'NO' WHEN '1' THEN 'YES' END 'List No internet' FROM [ACTIVITY_PACKAGE_ACTIVITIES] apa JOIN [ACTIVITY_PACKAGES] ap ON apa.ACTIVITY_PACKAGE_ID =ap.[ACTIVITY_PACKAGE_ID] JOIN [ACTIVITIES] a1 ON ap.ACTIVITY_ID = a1.ACTIVITY_ID JOIN [SITES] s ON a1.[SITE_ID] = s.[SITE_ID] JOIN [ACTIVITIES] a2 ON apa.[ACTIVITY_ID] = a2.ACTIVITY_ID WHERE a1.SEASON_ID = {season_id} ORDER BY S.SITENAME",
    }

    for name, query in queries.items():
       run_query(browser, name, query)
    

if __name__ == '__main__':
    main()