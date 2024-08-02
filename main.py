import os
import time
from io import BytesIO
from PIL import Image
from fpdf import FPDF
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.microsoft import EdgeChromiumDriverManager

WAIT_TIME = 1

def setup_driver():
    # 配置 Edge 瀏覽器選項
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1920x1080")
    
    # 初始化 WebDriver
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
    return driver

def capture_screenshots(driver, url):
    driver.get(url)
    time.sleep(WAIT_TIME)  # 等待網頁加載

    # 點擊進入全屏模式
    fullscreen_button = driver.find_element(By.CLASS_NAME, 'fullscreen-button')
    ActionChains(driver).move_to_element(fullscreen_button).click().perform()
    time.sleep(WAIT_TIME)

    page = 1
    screenshots = []

    while True:
        # 截圖並保存
        element = driver.find_element(By.CLASS_NAME, 'backgrounds')
        location = element.location
        size = element.size
        png = driver.get_screenshot_as_png()
        
        with Image.open(BytesIO(png)) as im:
            left = location['x']
            top = location['y']
            right = location['x'] + size['width']
            bottom = location['y'] + size['height']
            im = im.crop((left, top, right, bottom))
            
            screenshot_filename = f'screenshot_{str(page).zfill(2)}.png'
            im.save(screenshot_filename)
            screenshots.append(screenshot_filename)
            print(f'Saved {screenshot_filename}')

        # 先嘗試向下翻頁
        try:
            down = driver.find_element(By.CLASS_NAME, 'navigate-down')
            if down.is_enabled():
                ActionChains(driver).move_to_element(down).click().perform()
                time.sleep(WAIT_TIME)  # 等待翻頁動畫
                page += 1
                continue
        except Exception as e:
            print(f"Error navigating down: {e}")

        # 如果不能再向下，嘗試向右翻頁
        try:
            right = driver.find_element(By.CLASS_NAME, 'navigate-right')
            if right.is_enabled():
                ActionChains(driver).move_to_element(right).click().perform()
                time.sleep(WAIT_TIME)  # 等待翻頁動畫
                page += 1
                continue
        except Exception as e:
            print(f"Error navigating right: {e}")

        # 如果兩者都不能進行，則結束
        break

    return screenshots

def create_pdf(screenshots, output_filename):
    if not screenshots:
        print("No screenshots to create a PDF.")
        return

    first_image = Image.open(screenshots[0])
    width, height = first_image.size
    first_image.close()

    pdf = FPDF('L', 'mm', (height, width))
    pdf.set_margins(0, 0, 0)

    for image in screenshots:
        pdf.add_page()
        pdf.image(image, y=0, w=width)
    
    pdf.output(output_filename, 'F')
    print(f'PDF created: {output_filename}')

    # 清理截圖
    for image in screenshots:
        os.remove(image)

def get_page_title(driver):
    # 獲取當前頁面的標題並作為文件名
    title = driver.title
    return title.replace(' ', '_').replace(':', '').replace('/', '_')

def process_presentations(file_path):
    driver = setup_driver()
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)  # 確保輸出目錄存在

    with open(file_path, 'r') as file:
        urls = file.readlines()

    for url in urls:
        url = url.strip()
        if not url:
            continue

        try:
            print(f"Processing: {url}")
            screenshots = capture_screenshots(driver, url)
            title = get_page_title(driver)
            output_filename = os.path.join(output_dir, f'{title}.pdf')
            create_pdf(screenshots, output_filename)
        except Exception as e:
            print(f"Failed to process {url}: {e}")

    driver.quit()

if __name__ == '__main__':
    file_path = 'slides_urls.txt'  # 包含所有 Slides.com 簡報 URL 的文本文件
    process_presentations(file_path)


