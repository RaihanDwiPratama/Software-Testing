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
    
    # Mengklik menu my info
    submenu_myinfo = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[6]/a/span')
    actions = ActionChains(driver)
    actions.move_to_element(submenu_myinfo).perform()
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[6]/a/span').click()

    time.sleep(3)
    
    #--------------------------------------- 1

    ## Mengedit data tanpa memasukan data apapun
    
    time.sleep(2)
    
    # Submit data
    submit_info = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[5]/button')
    submit_info.click()

    time.sleep(3)

    #--------------------------------------- 2

    ## Mengedit data dengan memasukan beberapa data
    
    # Input lisence expiry date
    lisence_expriy_date_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[2]/div[2]/div/div[2]/div/div/input')))
    lisence_expriy_date_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    lisence_expriy_date_input.send_keys('2028-01-02')
    
    # Input SSN Number
    ssn_number_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[3]/div[1]/div/div[2]/input')))
    ssn_number_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    ssn_number_input.send_keys('3002')
    
    # Input SIN Number
    sin_number_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[3]/div[2]/div/div[2]/input')))
    sin_number_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    sin_number_input.send_keys('3729')
    
    time.sleep(3)
    
    # Submit data
    submit_info = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[5]/button')
    submit_info.click()

    time.sleep(3)
    
    #--------------------------------------- 3

    ## Mengedit data input "id" atau "Number" dengan memasukan data alfabet bukan numberic
    
    # Input Employee id
    employee_id_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div[2]/div/div/div[2]/input')))
    employee_id_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    employee_id_input.send_keys('asia')
    
    # Input Other id
    employee_id_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[1]/div[2]/div/div[2]/input')))
    employee_id_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    employee_id_input.send_keys('south')

    # Input Driver lisence
    driver_lisence_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[2]/div[1]/div/div[2]/input')))
    driver_lisence_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    driver_lisence_input.send_keys('nice')
    
    # Input SSN Number
    ssn_number_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[3]/div[1]/div/div[2]/input')))
    ssn_number_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    ssn_number_input.send_keys('one')
    
    # Input SIN Number
    sin_number_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[3]/div[2]/div/div[2]/input')))
    sin_number_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    sin_number_input.send_keys('three')
    
    # Input Military Service
    military_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[4]/div/div[1]/div/div[2]/input')))
    military_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    military_input.send_keys('military')
    
    time.sleep(3)
    
    # Submit data
    submit_info = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[5]/button')
    submit_info.click()

    time.sleep(3)

    #--------------------------------------- 4

    ## Mengedit data dengan memasukan semua data yang benar
    
    time.sleep(2)
    
    # Input nama
    first_name_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div/div/div[2]/div[1]/div[2]/input')))
    first_name_input.send_keys(Keys.COMMAND, 'a')  
    first_name_input.send_keys(Keys.DELETE)

    middle_name_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div/div/div[2]/div[2]/div[2]/input')))
    middle_name_input.send_keys(Keys.COMMAND, 'a')
    middle_name_input.send_keys(Keys.DELETE)

    last_name_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div/div/div/div[2]/div[3]/div[2]/input')))
    last_name_input.send_keys(Keys.COMMAND, 'a')
    last_name_input.send_keys(Keys.DELETE)
    # Mengisi data baru
    first_name_input.send_keys('Steves')
    middle_name_input.send_keys('Derve')
    last_name_input.send_keys('Smith')

    # Input Nickname
    nickname_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div[2]/div/div/div[2]/input')))
    nickname_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    nickname_input.send_keys('StevesD')

    # Input Employee id
    employee_id_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[1]/div[2]/div/div/div[2]/input')))
    employee_id_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    employee_id_input.send_keys('0021')
    
    # Input Other id
    employee_id_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[1]/div[2]/div/div[2]/input')))
    employee_id_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    employee_id_input.send_keys('3114')

    # Input Driver lisence
    driver_lisence_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[2]/div[1]/div/div[2]/input')))
    driver_lisence_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    driver_lisence_input.send_keys('1223')

    # Input lisence expiry date
    lisence_expriy_date_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[2]/div[2]/div/div[2]/div/div/input')))
    lisence_expriy_date_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    lisence_expriy_date_input.send_keys('2028-01-01')
    
    # Input SSN Number
    ssn_number_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[3]/div[1]/div/div[2]/input')))
    ssn_number_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    ssn_number_input.send_keys('3004')
    
    # Input SIN Number
    sin_number_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[2]/div[3]/div[2]/div/div[2]/input')))
    sin_number_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    sin_number_input.send_keys('3720')
    
    # Input Date birth
    birth_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[2]/div[1]/div/div[2]/div/div/input')))
    birth_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    birth_input.send_keys('1957-10-11')
    
    # Input Military Service
    military_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[4]/div/div[1]/div/div[2]/input')))
    military_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    military_input.send_keys('1222')

    # Nationality
    nationality_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[1]/div[1]/div/div[2]/div/div[1]/div[1]')
    nationality_input.click()
    time.sleep(1)
    nationality_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[1]/div[1]/div/div[2]/div/div[1]/div[1]')
    nationality_option.send_keys('Omani')
    nationality_option.send_keys(Keys.ENTER)
    
    # Martial status
    status_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[1]/div[2]/div/div[2]/div/div/div[1]')
    status_input.click()
    time.sleep(1)
    status_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[3]/div[1]/div[2]/div/div[2]/div/div/div[1]')
    status_option.send_keys('Single')
    status_option.send_keys(Keys.ENTER)
    
    time.sleep(3)

    # Submit data
    submit_info = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/div/div[2]/div[1]/form/div[5]/button')
    submit_info.click()

    time.sleep(3)