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
    wait = WebDriverWait(driver, 2)
    time.sleep(3)  # Delay sejenak agar lebih mudah dilihat

    # Isi formulir login dengan username dan password yang benar
    driver.find_element(By.NAME, "username").send_keys('Admin')
    driver.find_element(By.NAME, "password").send_keys('admin123' + Keys.ENTER)

    time.sleep(3)
    
    # Mengklik menu time
    submenu_time = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[4]/a')
    actions = ActionChains(driver)
    actions.move_to_element(submenu_time).perform()
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[4]/a').click()

    time.sleep(3)
    
    # Mengklik menu project info > customer
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[4]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[4]/ul/li[1]'))).click()

    time.sleep(3)

    #--------------------------------------- 1
    
    ## Menambahkan customer tanpa mengisi apapun
    
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div[1]/div/button').click()
    
    # Submit data
    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]')))
    submit_button.click()
    
    time.sleep(3)
    
    #--------------------------------------- 2
    
    ## Menambahkan customer dengan mengisi deskripsi saja
    driver.refresh()
    time.sleep(1)
    input_desc = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/textarea')))
    input_desc.send_keys('A PepsiCo associate innovating water-saving techniques ... water conservation efforts that make the company more sustainable and improve local communities.')
    
    # Submit data
    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]')))
    submit_button.click()
    
    time.sleep(3)
    
    #--------------------------------------- 3
    
    ## Mengisi data customer dengan full data, lalu di cancel dan masuk lagi ke menu add apakah data sebelumnya tersimpan atau tidak
    driver.refresh()
    time.sleep(1)
    input_name = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[2]/input')))
    input_name.send_keys('Pepsi')

    input_desc = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/textarea')))
    input_desc.send_keys('A PepsiCo associate innovating water-saving techniques ... water conservation efforts that make the company more sustainable and improve local communities.')
    
    # Cancel data
    cancel_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[1]')))
    cancel_button.click()
    
    # Mengklik menu project info > customer > add
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[4]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[4]/ul/li[1]'))).click()
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div[1]/div/button'))).click()

    time.sleep(3)
    
    #--------------------------------------- 4
    
    ## Menambahkan customer dengan mengisi nama saja
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[2]/input').send_keys('Pepsi')
    
    # Submit data
    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]')))
    submit_button.click()
    
    time.sleep(3)
    
    #--------------------------------------- 5
    
    ## Menambahkan customer dengan full data yang benar
    
    # Mengklik menu project info > customer > add
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[4]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[2]/nav/ul/li[4]/ul/li[1]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div[1]/div/button'))).click()
    time.sleep(3)
    
    input_name = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[2]/input')))
    input_name.send_keys('Sprite')
    input_desc = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/textarea')))
    input_desc.send_keys('A PepsiCo associate innovating water-saving techniques ... water conservation efforts that make the company more sustainable and improve local communities.')
    
    # Submit data
    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]')))
    submit_button.click()
    
    time.sleep(3)
    
    #---------------------------------------
    
    