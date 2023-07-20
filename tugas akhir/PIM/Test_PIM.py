import pytest
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook

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

def test_pim(driver):
    # Kasus Pengujian Positive
    time.sleep(3)  # Delay sejenak agar lebih mudah dilihat

    # Isi formulir login dengan username dan password yang benar
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[1]/div/div[2]/input').send_keys('Admin')  # Menggunakan By.NAME untuk input username
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[2]/div/div[2]/input').send_keys('admin123' + Keys.ENTER)  # Menggunakan By.NAME untuk input password dan Keys.ENTER untuk login

    time.sleep(3)
    
    # Kasus Pengujian Positive
    # Mengklik submenu "PIM"
    submenu_apply = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[2]/a')
    submenu_apply.click()
    
    time.sleep(5)
    
    # Cari karyawan dengan ID
    employee_id = '0038'
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/input').send_keys(employee_id)
    
    # Cek apakah hasil pencarian sesuai dengan ID karyawan
    driver.find_element(By.XPATH,'//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]').click()
    time.sleep(5)
    
    # Verifikasi hasil pencarian
    search_result_id = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[2]/div').text
    assert search_result_id == employee_id, 'Search result mismatch - Positive Test'
    time.sleep(5)
    
    # Kasus Pengujian Negative
    # Mengklik submenu "PIM"
    submenu_apply = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[2]/a')
    submenu_apply.click()
    
    time.sleep(5)
    
    # Cari karyawan dengan ID
    employee_id = '0101'
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/input').send_keys(employee_id)
    
    # Cek apakah hasil pencarian sesuai dengan ID karyawan
    driver.find_element(By.XPATH,'//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]').click()
    time.sleep(5)
    
    # Verifikasi hasil pencarian negatif
    search_result_id = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[41]/div').text
    assert search_result_id == employee_id, 'Search result mismatch - Negative Test'
    