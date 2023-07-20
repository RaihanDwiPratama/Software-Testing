import pytest
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException

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

# Kasus Pengujian Positive
def test_valid_username_and_password(driver):
    try:
        time.sleep(3)  # Delay sejenak agar lebih mudah dilihat

        # Isi formulir login dengan username dan password yang benar 
        driver.find_element(By.NAME, "username").send_keys('Admin')

        driver.find_element(By.NAME, "password").send_keys('admin123' + Keys.ENTER)

        # time.sleep(3)# Submit formulir login dengan tombol ENTER
        # send_keys(Keys.ENTER)

        # Tunggu halaman setelah login dimuat
        time.sleep(5)  # Delay lebih lama agar halaman terbuka sepenuhnya

        # Verifikasi apakah login berhasil
        welcome_message = driver.find_element(By.ID, "welcome")
        assert welcome_message.text == "Welcome Admin"
        print("Test Valid Username and Password: Passed") 

    except NoSuchElementException:
        print("Test Valid Username and Password: Failed") 


# Kasus Pengujian Negative
def test_empty_username_and_password(driver):
    try:
        time.sleep(3)  # Delay sejenak agar lebih mudah dilihat

        # # Submit formulir login tanpa mengisi username dan password
        # driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[3]/button').click()
        # Submit formulir login tanpa mengisi username dan password
        login_button = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[3]/button').send_keys(Keys.ENTER)

        time.sleep(5)  # Delay sejenak agar lebih mudah dilihat

        # Verifikasi apakah ada pesan kesalahan
        error_message = driver.find_element(By.ID, "spanMessage")
        assert error_message.text == "Username cannot be empty"
        print("Test Empty Username and Password: Passed")

    except NoSuchElementException:
        print("Test Empty Username and Password: Failed")

def test_invalid_username_and_valid_password(driver):
    try:
        time.sleep(2)  # Delay sejenak agar lebih mudah dilihat
        
        # Isi formulir login dengan username yang salah dan password yang benar
        driver.find_element(By.NAME, "username").send_keys('invalidusername')

        driver.find_element(By.NAME, "password").send_keys('admin123' + Keys.ENTER)

        time.sleep(5)  # Delay sejenak agar lebih mudah dilihat

        # Verifikasi apakah ada pesan kesalahan
        error_message = driver.find_element(By.ID, "spanMessage")
        assert error_message.text == "Invalid credentials"
        print("Test Invalid Username and Valid Password: Passed")

    except NoSuchElementException:
        print("Test Invalid Username and Valid Password: Failed")

def test_valid_username_and_invalid_password(driver):
    try:
        time.sleep(2)  # Delay sejenak agar lebih mudah dilihat
        
        # Isi formulir login dengan username yang benar dan password yang salah
        driver.find_element(By.NAME, "username").send_keys('Admin')

        driver.find_element(By.NAME, "password").send_keys('invalidpassword' + Keys.ENTER)

        time.sleep(5)  # Delay sejenak agar lebih mudah dilihat

        # Verifikasi apakah ada pesan kesalahan
        error_message = driver.find_element(By.ID, "spanMessage")
        assert error_message.text == "Invalid credentials"
        print("Test Valid Username and Invalid Password: Passed")

    except NoSuchElementException:
        print("Test Valid Username and Invalid Password: Failed")
