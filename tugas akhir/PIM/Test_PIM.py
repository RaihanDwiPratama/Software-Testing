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
    
    # Mengklik submenu "PIM"
    submenu_apply = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[2]/a')
    submenu_apply.click()
    
    time.sleep(5)
    # Cari karyawan dengan ID
    employee_id = '0015'
    driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[1]/div/div[2]/div/div[2]/input').send_keys(employee_id)
    
    # Cek apakah hasil pencarian sesuai dengan ID karyawan
    driver.find_element(By.XPATH,'//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]').click()
    time.sleep(5)

    # Membuat workbook dan worksheet baru
    workbook = Workbook()
    worksheet = workbook.active

    # Mendapatkan data yang ingin disimpan
    data1 = "Data 1"
    data2 = "Data 2"

    # Menyimpan data ke dalam worksheet
    worksheet['A1'] = 'Data 1'
    worksheet['B1'] = 'Data 2'
    worksheet['A2'] = data1
    worksheet['B2'] = data2

    # Menyimpan workbook ke dalam file Excel
    workbook.save('hasil_testing.xlsx')



    # # Mengklik menu leave
    # submenu_leave = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/aside/nav/div[2]/ul/li[2]/a')
    # actions = ActionChains(driver)
    # actions.move_to_element(submenu_leave).perform()
    # driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[3]/button').click()

    # time.sleep(10)

    # Navigasi ke halaman View Employee
    # driver.get('https://opensource-demo.orangehrmlive.com/web/index.php/pim/viewEmployee')

    # # Cek apakah halaman View Employee berhasil dibuka
    # page_title = driver.title
    # assert page_title == 'OrangeHRM', 'Failed to open View Employee page'

    # # Cari karyawan dengan ID
    # employee_id = '0015'
    # search_box = driver.find_element_by_id('empsearch_id')
    # search_box.clear()
    # search_box.send_keys(employee_id)
    # driver.find_element(By.XPATH,'//*[@id="app"]/div[1]/div[2]/div[2]/div/div[1]/div[2]/form/div[2]/button[2]').click()

    # # Cek apakah hasil pencarian sesuai dengan ID karyawan
    # search_result_id = driver.find_element_by_xpath("//table[@id='resultTable']/tbody/tr/td[2]").text
    # assert search_result_id == employee_id, 'Search result mismatch'



# from selenium import webdriver

# # Inisialisasi objek Chrome WebDriver
# driver = webdriver.Chrome()

# # Buka URL
# driver.get('https://opensource-demo.orangehrmlive.com')

# # Masukkan username dan password
# driver.find_element_by_xpath('//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[1]/div/div[2]').send_keys('admin')
# driver.find_element_by_xpath('//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[2]/div/div[2]').send_keys('admin123')

# # Klik tombol Login
# driver.find_element_by_xpath('//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[3]/button').click()

# # Cek apakah login berhasil
# welcome_message = driver.find_element_by_id('welcome').text
# assert welcome_message == 'Welcome Admin', 'Login failed'

# # Navigasi ke halaman View Employee
# driver.get('https://opensource-demo.orangehrmlive.com/web/index.php/pim/viewEmployee')

# # Cek apakah halaman View Employee berhasil dibuka
# page_title = driver.title
# assert page_title == 'OrangeHRM', 'Failed to open View Employee page'

# # Cari karyawan dengan ID
# employee_id = '0015'
# search_box = driver.find_element_by_id('empsearch_id')
# search_box.clear()
# search_box.send_keys(employee_id)
# driver.find_element_by_id('searchBtn').click()

# # Cek apakah hasil pencarian sesuai dengan ID karyawan
# search_result_id = driver.find_element_by_xpath("//table[@id='resultTable']/tbody/tr/td[2]").text
# assert search_result_id == employee_id, 'Search result mismatch'

# # Tutup WebDriver
# driver.quit()
