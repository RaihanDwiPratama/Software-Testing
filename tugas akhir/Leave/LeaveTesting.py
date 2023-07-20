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
    
    # Mengklik menu leave
    submenu_leave = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[3]/a')
    actions = ActionChains(driver)
    actions.move_to_element(submenu_leave).perform()
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[3]/a').click()

    time.sleep(3)
     
    # Mengklik menu Report > Leave Entitlments
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[4]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[4]/ul/li[1]'))).click()
    time.sleep(3)

    #--------------------------------------- 1
    
    ## Mencari dengan data yang kosong

    # Generate pencarian
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div[2]/form/div[3]/button')
    search .click()

    time.sleep(3)

    #--------------------------------------- 2
    
    ## Mencari dengan salah 1 data

    # Sub unit
    sub_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/div/div[4]/div/div[2]/div/div/div[1]')))
    sub_input.click()
    time.sleep(1)
    sub_input.send_keys('TechOps')
    time.sleep(1)
    sub_input.send_keys(Keys.ENTER)

    # Generate pencarian
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div[2]/form/div[3]/button')
    search .click()

    time.sleep(3)

    #--------------------------------------- 3
    
    ## Mencari dengan data yang benar

    driver.refresh()

    # Leave type
    # Mencari elemen dropdown
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div[2]/form/div[2]/div/div[1]/div/div[2]/div/div/div[1]')))
    # Klik elemen dropdown
    actions = ActionChains(driver)
    actions.move_to_element(dropdown).click().perform()
    # Tunggu hingga dropdown sepenuhnya terbuka
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div[2]/form/div[2]/div/div[1]/div/div[2]/div/div/div[2]')))
    # Cari dan pilih opsi dengan mengirimkan teks dan menekan tombol ENTER
    dropdown.send_keys('Test')
    dropdown.send_keys(Keys.ENTER)

    # Location
    location_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/div/div[3]/div/div[2]/div/div/div[1]')))
    location_input.click()
    time.sleep(1)
    location_input.send_keys('canadian')
    time.sleep(1)
    location_input.send_keys(Keys.ENTER)

    # Sub unit
    sub_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/div/div[4]/div/div[2]/div/div/div[1]')))
    sub_input.click()
    time.sleep(1)
    sub_input.send_keys('TechOps')
    time.sleep(1)
    sub_input.send_keys(Keys.ENTER)
    
    # Job title
    job_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/div/div[5]/div/div[2]/div/div/div[1]')
    job_input.click()
    time.sleep(1)
    job_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/div/div[5]/div/div[2]/div/div/div[1]')
    job_option.send_keys('HR')
    job_option.send_keys(Keys.ENTER)

    time.sleep(3)  
    
    # Generate pencarian
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div[2]/form/div[3]/button')
    search .click()

    time.sleep(3)
