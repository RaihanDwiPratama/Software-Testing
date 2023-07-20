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
    
    #---------------------------------------

    ## Tes semua button yang ada
    
    time.sleep(2)
    # Time at work
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/div[1]/div[2]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()

    time.sleep(2)
    # My action > leave
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div[1]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(2)
    # My action > timesheets
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div[2]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(2)
    # My action > pending
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div[3]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(2)
    # Quick Launch > assign leave
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[3]/div/div[2]/div/div[1]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(2)
    # Quick Launch > leave list
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[3]/div/div[2]/div/div[2]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(2)
    # Quick Launch > timesheets
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[3]/div/div[2]/div/div[3]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(2)
    # Quick Launch > Apply leaves
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[3]/div/div[2]/div/div[4]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(2)
    # Quick Launch > My leaves
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[3]/div/div[2]/div/div[5]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(2)
    # Quick Launch > My timesheets
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[3]/div/div[2]/div/div[6]/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[8]/a'))).click()
    
    time.sleep(3)