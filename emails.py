import customtkinter as ctk
from tkinter import messagebox
from PIL import Image
import openpyxl
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pyautogui
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import os
import re
import pdfplumber
import pandas as pd
import customtkinter as ctk
import json
import threading
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import tkinter as tk
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def enviar_emails():
    driver = webdriver.Chrome()
    driver.get('https://www.mautic.org/')

    botao_register = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/user/login')]"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", botao_register)
    driver.execute_script("arguments[0].click();", botao_register)


    login = 'luiz.logika@gmail.com'
    password = 'Dev123@'
    logar = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@class='auth0-lock-input']"))
    )
    logar.send_keys(login)

    password_enter = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='1-password']"))
    )
    password_enter.send_keys(password)

    sleep(5)
    driver.quit()

enviar_emails()