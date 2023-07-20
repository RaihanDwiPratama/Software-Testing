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
    submenu_performance = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[7]/a')
    actions = ActionChains(driver)
    actions.move_to_element(submenu_performance).perform()
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[7]/a').click()

    time.sleep(3)
     
    # Mengklik menu Report > Leave Entitlments
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[1]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[1]/ul/li[1]'))).click()
    time.sleep(2)

    #--------------------------------------- 1

    ## Mencari dengan data yang kosong

    # Generate pencarian
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]')
    search.click()

    time.sleep(3)

    #--------------------------------------- 2

    ## Reset data dengan data yang kosong

    # Reset data
    reset = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[1]')
    reset.click()

    time.sleep(3)

    #--------------------------------------- 3

    ## Mencari dengan data yang benar

    # Job title
    job_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div/div/div[2]/div/div/div[1]')
    job_input.click()
    time.sleep(1)
    job_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div/div/div[2]/div/div/div[1]')
    job_option.send_keys('Content')
    job_option.send_keys(Keys.ENTER)

    # Generate pencarian
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]')
    search.click()

    time.sleep(3)  

    #--------------------------------------- 4

    ## Reset data dengan data yang tersedia sebelumnya

    # Reset data
    reset = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[1]')
    reset.click()

    time.sleep(3)

    #--------------------------------------- 5

    # Menambahkan data dengan data yang kosong

    # Add
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[2]/div[1]/button'))).click()

    time.sleep(2)

    # Save data
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]'))).click()

    time.sleep(3)

    #--------------------------------------- 6

    # Menambahkan semua data dengan data yang benar
    driver.refresh()
    time.sleep(2)

    # key performance
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[1]/div/div[2]/input'))).send_keys('Dilevery')

    # Job title
    job_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[2]/div/div[2]/div/div/div[1]')
    job_input.click()
    time.sleep(1)
    job_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[2]/div/div[2]/div/div/div[1]')
    job_option.send_keys('Content')
    job_option.send_keys(Keys.ENTER)

    # Minimum rating
    rating_min_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')))
    rating_min_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    rating_min_input.send_keys('0')

    # Max rating
    rating_max_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')))
    rating_max_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    rating_max_input.send_keys('10')

    time.sleep(2)

    # Save data
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]'))).click()

    time.sleep(3)



    