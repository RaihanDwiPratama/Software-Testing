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
    time.sleep(3)  # Delay sejenak agar lebih mudah dilihat

    # Isi formulir login dengan username dan password yang benar
    driver.find_element(By.NAME, "username").send_keys('Admin')
    driver.find_element(By.NAME, "password").send_keys('admin123' + Keys.ENTER)

    time.sleep(3)
    
    # Mengklik menu admin
    submenu_admin = driver.find_element(By.CSS_SELECTOR, 'a.oxd-main-menu-item')
    actions = ActionChains(driver)
    actions.move_to_element(submenu_admin).perform()
    driver.find_element(By.CSS_SELECTOR, 'a.oxd-main-menu-item').click()

    time.sleep(3)
    
    #--------------------------------------- 1
    
    ## Input username, userrole, employee name, status yang benar >> bug error pada employee name
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[1]/div/div[2]/input').send_keys('Admin')
    
    user_role_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/div/div/div[1]')
    user_role_input.click()
    time.sleep(1)
    user_role_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/div/div/div[1]')
    user_role_option.send_keys('admin')
    user_role_option.send_keys(Keys.ENTER)
    
    # driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[3]/div/div[2]/div/div/input').send_keys('Maceo Test Plex')
    
    status_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[4]/div/div[2]/div/div/div[1]')
    status_input.click()
    status_input.send_keys('Enable') 
    status_input.send_keys(Keys.ENTER)
    
    time.sleep(3)
    
    # Submit data
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]')
    search.click()
    time.sleep(5)
    
    # Reset data
    reset = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[1]')
    reset.click()
    time.sleep(5)
    
    #--------------------------------------- 2
    
    ## Input username, employee name yang salah >> ketika mengisi user role yang salah, maka tidak bisa melakukan submit, dan tidak menampilkan alert data tidak ada yang sama ketika username yg tidak tersedia datanya
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[1]/div/div[2]/input').send_keys('Admini')
    
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[3]/div/div[2]/div/div/input').send_keys('Ads')
    
    time.sleep(3)
    
    # Submit data
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]')
    search.click()
    time.sleep(5)
    
    # Reset data
    reset = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[1]')
    reset.click()
    time.sleep(5)
    
    #--------------------------------------- 3
    
    ## Input kosong lalu tetap disubmit lalu di reset data walau kosong
    
    # Submit data
    search = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]')
    search.click()
    time.sleep(5)
    
    # Reset data
    reset = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[1]')
    reset.click()
    time.sleep(5)
    
    #--------------------------------------- 4
    
    ## Menambah user baru tanpa input data apapun
    add_new_user = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[2]/div[1]/button')
    add_new_user.click()
    
    time.sleep(3)
    
    # Submit data
    submit_user = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]')
    submit_user.click()
    time.sleep(5)

    #--------------------------------------- 5
    
    # Memasukan password yang tidak match dengan confirm password
    driver.refresh()
    time.sleep(3)
    
    password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')
    password_input.send_keys('Admini123')
    
    confirm_password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')
    confirm_password_input.send_keys('Admini1234')
    
    time.sleep(3)
    
    #--------------------------------------- 6
    
    ## Memasukan password yang sama dengan confirm password tetapi tidak terdapat angka
    driver.refresh()
    time.sleep(3)
    
    password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')
    password_input.send_keys('AADDMMIIN')
    
    confirm_password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')
    confirm_password_input.send_keys('AADDMMIIN')
    
    time.sleep(3)
    
    #--------------------------------------- 7
    
    ## Memasukan password yang sama dengan confirm password tetapi < 7 karakter dan tidak ada angka
    driver.refresh()
    time.sleep(3)
    
    password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')
    password_input.send_keys('admini')
    
    confirm_password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')
    confirm_password_input.send_keys('admini')
    
    time.sleep(3)
    
    #--------------------------------------- 8
    
    ## Memasukan password yang sama dengan confirm password tetapi < 7 karakter dan terdapat angka
    driver.refresh()
    time.sleep(3)
    
    password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')
    password_input.send_keys('admini1')
    
    confirm_password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')
    confirm_password_input.send_keys('admini1')
    
    time.sleep(3)
    
    ## Memasukan password yang sama dengan confirm password tetapi > 64 karakter
    driver.refresh()
    time.sleep(3)
    
    password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')
    password_input.send_keys('adminiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii123')
    
    confirm_password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')
    confirm_password_input.send_keys('adminiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii123')
    
    time.sleep(3)
    
    ## Memasukan password yang sama dengan confirm password tetapi > 64 karakter dan tidak ada angka
    driver.refresh()
    time.sleep(3)
    
    password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')
    password_input.send_keys('adminiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii')
    
    confirm_password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')
    confirm_password_input.send_keys('adminiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii')
    
    time.sleep(3)
    
    #--------------------------------------- 9
    
    ## Memasukan password yang berbeda dengan confirm password tetapi > 64 karakter dan tidak ada angka
    driver.refresh()
    time.sleep(3)
    
    password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')
    password_input.send_keys('adminiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii')
    
    confirm_password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')
    confirm_password_input.send_keys('adminiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii')
    
    time.sleep(3)
    
    # --------------------------------------- 10
    
    # Menambah user baru >> employee name harus di klik dropdownnya, jika tidak maka dianggap invalid padahal data sudah benar
    driver.refresh()
    
    time.sleep(3)
    
    user_role_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[1]/div/div[2]/div/div/div[1]')
    user_role_input.click()
    time.sleep(1)
    user_role_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[1]/div/div[2]/div/div/div[1]')
    user_role_option.send_keys('admin')
    user_role_option.send_keys(Keys.ENTER)
    
    status_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[3]/div/div[2]/div/div/div[1]')
    status_input.click()
    time.sleep(1)
    status_option = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[3]/div/div[2]/div/div/div[1]')
    status_option.send_keys('enabled')
    status_option.send_keys(Keys.ENTER)
    
    password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[1]/div/div[2]/input')
    password_input.send_keys('Admini123')
    
    confirm_password_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[2]/div/div[2]/div/div[2]/input')
    confirm_password_input.send_keys('Admini123')
    
    # Memasukkan employee name
    employee_name_input = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[2]/div/div[2]/div/div/input')
    employee_name_input.click()

    # Klik dropdown menggunakan JavaScript
    dropdown = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[2]/div/div[2]/div')
    driver.execute_script("arguments[0].click();", dropdown)

    # Pilih data "Aaliyah Haq"
    actions.move_to_element(dropdown).send_keys('Aaliyah Haq').perform()
    time.sleep(1)
    actions.send_keys(Keys.ENTER).perform()
    
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[1]/div/div[4]/div/div[2]/input').send_keys('Admini')
    
    time.sleep(3)
    
    # Submit data
    submit_user = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]')
    submit_user.click()
    time.sleep(5)
    