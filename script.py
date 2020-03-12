from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.chrome import service
import time
from pandas import ExcelWriter
import pandas as pd
from requied_details import *

driver = webdriver.Chrome()
driver.get("http://www.linkedin.com")
# assert "Python" in driver.title
time.sleep(2)

email = driver.find_element_by_name("session_key")
email.send_keys(LINKEDIN_EMAIL)
password = driver.find_element_by_name("session_password")
password.send_keys(LINKEDIN_PASSWORD)
password.send_keys(Keys.RETURN)

time.sleep(2)

data = []

for i in range(1, PAGE_COUNT+1):
    driver.get(f'https://www.linkedin.com/search/results/people/?facetGeoRegion=%5B"us%3A0"%5D&keywords=CEO&origin=FACETED_SEARCH&page={i}')

    for i in range(1, 1000, 15):
        driver.execute_script(f"window.scrollTo(0, {i})") 

    time.sleep(5)

    elems = driver.find_elements_by_css_selector('.search-result__info .search-result__result-link')

    lst = [x.get_attribute('href') for x in elems ]

    for x in elems:
        href = x.get_attribute('href')
        if href.find('https://www.linkedin.com/in/') != -1:
            data.append(x.get_attribute('href'))

# df = pd.DataFrame({
#     'href': data,
# })
# # Create a Pandas Excel writer using XlsxWriter as the engine.
# writer = pd.ExcelWriter('linkedin_ceo_href.xlsx', engine='xlsxwriter')
# # Convert the dataframe to an XlsxWriter Excel object.
# df.to_excel(writer, sheet_name='Sheet1')
# # Close the Pandas Excel writer and output the Excel file.
# writer.save()


names, emails, phone_numbers, locations, current_positions, exps, companies, abouts, academies, degrees, interests, twitters, websites = ([] for i in range(13))

for link in data:
    driver.get(link)
    
    time.sleep(3)

    for i in range(1, 1500, 15):
        driver.execute_script(f"window.scrollTo(0, {i})")

    time.sleep(5)

    # Getting Experience
    experiences = driver.find_elements_by_css_selector('.pv-entity__position-group-pager')

    for experience in experiences:
        company_name, current_position, exp, about = ('N/A' for i in range(4))
        try:
            current_position = experience.find_elements_by_css_selector('.t-14.t-black.t-bold span')[1].text
            company_name = experience.find_elements_by_css_selector('.t-16.t-black.t-bold span')[1].text
            exp = experience.find_elements_by_css_selector('.t-14.t-black.t-normal span')[1].text
            
            print(f'{company_name} | {current_position} | {exp}')
            
        except Exception as e:
            # print(e)
            try:
                company_name = experience.find_element_by_css_selector('.pv-entity__secondary-title.t-14.t-black.t-normal').text
                current_position = experience.find_element_by_css_selector('.t-16.t-black.t-bold').text
                exp = experience.find_element_by_css_selector('.pv-entity__bullet-item-v2').text
                print(f'{company_name} | {current_position} |  {exp}')
            except Exception as ef:
                print(ef)
        
        break

    # Getting Education
    educations = driver.find_elements_by_css_selector('.pv-profile-section__list-item')
    
    tmp_academy, tmp_degree, tmp_interest = [], [], []

    for education in educations:
        academy_name, degree_title, interest = ('N/A' for i in range(3))

        try:
            academy_name = education.find_element_by_css_selector('.pv-entity__degree-info h3').text
            degree_title = education.find_elements_by_css_selector('.pv-entity__degree-info p span')[1].text
            print(academy_name, degree_title)
        except Exception as e:
            print(e)

        tmp_academy.append(academy_name)
        tmp_degree.append(degree_title)

    # GETTING INTERESTS
    try:
        interest = driver.find_elements_by_css_selector('.pv-entity__summary-title-text')
        for inter in interest:
            tmp_interest.append(inter.text)

        tmp_interest = (' | '.join( x for x in tmp_interest if x != 'N/A'))
    except Exception as e:
        print(e)


    tmp_academy = (' | '.join(  x for x in tmp_academy if x != 'N/A'))
    tmp_degree = (' | '.join(x for x in tmp_degree if x != 'N/A'))
    
    academies.append(tmp_academy)
    degrees.append(tmp_degree)
    interests.append(tmp_interest)
    
    try:
        see_more_button = driver.find_element_by_css_selector('.lt-line-clamp__more')
        driver.execute_script("arguments[0].click();", see_more_button)
        about = driver.find_element_by_css_selector('.lt-line-clamp__raw-line').text
    except Exception as e:
        print(e)
    
    abouts.append(about)
    companies.append(company_name)
    current_positions.append(current_position)
    exps.append(exp)

    # Initializing required values
    name = driver.find_element_by_css_selector('.inline.t-24.t-black.t-normal.break-words').text
    email = 'N/A'
    phone_number = 'N/A'
    twitter = 'N/A'
    website = 'N/A'
    
    location = driver.find_element_by_css_selector('.t-16.t-black.t-normal.inline-block').text
    contry = 'N/A'

    # Changing URL to contact info
    contact_link = link + 'detail/contact-info/'
    driver.get(contact_link)
    
    # For getting Email Address:
    em = driver.find_elements_by_css_selector('.ci-email div a')
    
    for x in em:
        email = x.text
        break

    # For getting Phone Number:
    pn = driver.find_elements_by_css_selector('.ci-phone ul li .t-14.t-black.t-normal')
    for x in pn:
        phone_number = x.text
        break

    # For getting Twitter link:
    tw = driver.find_elements_by_css_selector('.ci-twitter ul li .t-14.t-black.t-normal')
    for x in tw:
        twitter = x.get_attribute("href")
        break

    # For getting Website link:
    web = driver.find_elements_by_css_selector('.ci-websites ul li div .t-14.t-black.t-normal')
    
    for x in web:
        website = x.get_attribute("href")
        break

    names.append(name)
    emails.append(email)
    phone_numbers.append(phone_number)
    locations.append(location)
    twitters.append(twitter)
    websites.append(website)

    # print(f'{name} | {email} | {phone_number} | {location}')

print(len(data))

df = pd.DataFrame({
    'Name': names,
    'Email': emails,
    'Phone Number': phone_numbers,
    'Twitter': twitters,
    'Website': websites,
    'Location': locations,
    'Current Position': current_positions,
    'Current Company': companies,
    'Experience': exps,
    'Education Institute': academies,
    'Education Degree': degrees,
    'Interests': interests,
    'About': abouts
    })

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter( f'{NAME_OF_OUTPUT_FILE}.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()

# assert "No results found." not in driver.page_source
driver.close()
