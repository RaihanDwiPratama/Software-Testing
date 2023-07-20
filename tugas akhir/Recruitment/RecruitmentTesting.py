import pytest
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

@pytest.fixture
def driver():
    # Inisialisasi WebDriver
    driver = webdriver.Chrome()
    # Buka halaman login
    driver.get("https://opensource-demo.orangehrmlive.com/")
    driver.maximize_window()
    yield driver
    # Tutup browser setelah pengujian selesai
    driver.quit()
    
def test(driver):
    wait = WebDriverWait(driver, 10)
    time.sleep(3)  # Delay sejenak agar lebih mudah dilihat

    # Isi formulir login dengan username dan password yang benar
    driver.find_element(By.NAME, "username").send_keys('Admin')
    driver.find_element(By.NAME, "password").send_keys('admin123' + Keys.ENTER)

    time.sleep(3)
    
    # Mengklik menu recruitment
    submenu_recruitment = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[5]/a/span')
    actions = ActionChains(driver)
    actions.move_to_element(submenu_recruitment).perform()
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[5]/a/span').click()

    time.sleep(3)
     
    # Mengklik menu Report > Leave Entitlments
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[2]'))).click()
    time.sleep(2)

    #--------------------------------------- 1

    ## Mencari dengan data yang benar

    # driver.refresh()
    time.sleep(2)
    
    # Job title
    job_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[1]/div/div[2]/div/div/div[1]')
    job_input.click()
    time.sleep(1)
    job_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[1]/div/div[2]/div/div/div[1]')
    time.sleep(1)
    job_option.send_keys('Chief')
    job_option.send_keys(Keys.ENTER)

    time.sleep(2)  

    # Vacancy
    vacancy_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/div/div/div[1]')
    vacancy_input.click()
    time.sleep(1)
    vacancy_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/div/div/div[1]')
    vacancy_option.send_keys('Sales')
    vacancy_option.send_keys(Keys.ENTER)

    time.sleep(2) 

    # Hiring Manager
    hiring_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[3]/div/div[2]/div/div/div[1]')
    hiring_input.click()
    time.sleep(1)
    hiring_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[3]/div/div[2]/div/div/div[1]')
    hiring_option.send_keys('Odis')
    hiring_option.send_keys(Keys.ENTER)

    # Status
    status_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[4]/div/div[2]/div/div/div[1]')
    status_input.click()
    time.sleep(1)
    status_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[4]/div/div[2]/div/div/div[1]')
    status_option.send_keys('Active')
    status_option.send_keys(Keys.ENTER)

    time.sleep(2) 

    # Generate pencarian
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]')
    search .click()

    time.sleep(3)


    #--------------------------------------- 2

    ## Mencari dengan data yang kosong

    driver.refresh()
    time.sleep(2)

    # Generate pencarian
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]')
    search .click()

    time.sleep(3)

    
    #--------------------------------------- 3

    ## Mencari dengan salah satu data

    driver.refresh()
    time.sleep(2)
    
    # Job title
    job_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[1]/div/div[2]/div/div/div[1]')
    job_input.click()
    time.sleep(1)
    job_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[1]/div/div[2]/div/div/div[1]')
    time.sleep(1)
    job_option.send_keys('Chief')
    job_option.send_keys(Keys.ENTER)

    time.sleep(2)  

    # Vacancy
    vacancy_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/div/div/div[1]')
    vacancy_input.click()
    time.sleep(1)
    vacancy_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/div/div/div[1]')
    vacancy_option.send_keys('Sales')
    vacancy_option.send_keys(Keys.ENTER)

    time.sleep(2) 

    # Generate pencarian
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]')
    search .click()

    time.sleep(3)
