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
    submenu_buzz = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[12]/a')
    actions = ActionChains(driver)
    actions.move_to_element(submenu_buzz).perform()
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[12]/a').click()

    time.sleep(3)

    #--------------------------------------- 1

    ## Filter dari like terbanyak
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/button[2]'))).click()

    time.sleep(3)
    
    #---------------------------------------

    ## Filter dari comment terbanyak
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/button[3]'))).click()

    time.sleep(3)

    #---------------------------------------

    ## Filter dari post terbaru
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/button[1]'))).click()

    time.sleep(3)
        
    #--------------------------------------- 2

    ## Post tanpa memasukan data
    
    # Post data
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[1]/div[1]/div[2]/form/div/div/button').click()
    
    time.sleep(4)
    
    #--------------------------------------- 3

    ## Post caption dan video
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[1]/div[1]/div[2]/form/div/textarea'))).send_keys('Hello! pop cat is waiting you!')
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[1]/div[2]/button[2]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/form/div[2]/div[2]/textarea'))).send_keys('https://youtu.be/eR8uA3kp8ec')
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/form/div[3]/button'))).click()

    time.sleep(3)

    # Post data
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[1]/div[1]/div[2]/form/div/div/button').click()

    time.sleep(3)

    #--------------------------------------- 4

    ## Love post
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="heart-svg"]'))).click()

    time.sleep(3)

    #---------------------------------------

    ## Comment post
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[3]/div[1]/button[1]'))).click()

    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[4]/div/form/div/div[2]/input'))).send_keys('waw' + Keys.ENTER)
    time.sleep(3)

    #--------------------------------------- 5

    ## Edit post dengan tidak mengubah data apapun
    driver.refresh()
    time.sleep(3)
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/ul/li[2]'))).click()

    time.sleep(2)

    # Post data
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div/div/form/div[4]/button').click()
    time.sleep(3)

    #--------------------------------------- 6

    ## Edit post dengan mengubah data lalu di close
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/ul/li[2]'))).click()

    caption_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div/div/form/div[1]/div[2]/div/textarea')))
    caption_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    caption_input.send_keys('Hello! pop cat is waiting you! haha')

    video_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div/div/form/div[2]/div[2]/textarea')))
    video_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    video_input.send_keys('https://youtube.com/shorts/_kKcKrj9Ce8?feature=share')

    time.sleep(3)

    # Close modal edit data
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div/div/button').click()
    time.sleep(3)

    #--------------------------------------- 7

    ## Edit post dengan mengubah data 
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/ul/li[2]'))).click()

    caption_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div/div/form/div[1]/div[2]/div/textarea')))
    caption_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    caption_input.send_keys('Hello! pop cat is waiting you! haha')

    video_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div/div/form/div[2]/div[2]/textarea')))
    video_input.send_keys(Keys.COMMAND, 'a', Keys.DELETE) 
    video_input.send_keys('https://youtube.com/shorts/_kKcKrj9Ce8?feature=share')

    time.sleep(3)

    # Post data
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/div/div/form/div[4]/button').click()
    time.sleep(3)

    #--------------------------------------- 8

    ## Delete post tapi di cancel
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/ul/li[1]/p'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[3]/div/div/div/div[3]/button[1]'))).click()

    time.sleep(3)

    #--------------------------------------- 9

    ## Delete post 
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/button'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div[2]/li/ul/li[1]/p'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[3]/div/div/div/div[3]/button[2]'))).click()

    time.sleep(3)