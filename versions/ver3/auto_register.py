import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime
import random

# Đường dẫn đến thư mục và file Excel
folder_path = r'C:\Users\MKT\Desktop\Elextrolux'
excel_path = os.path.join(folder_path, 'danh_sach_san_pham.xlsx')

# Số lượng tab muốn mở cùng lúc
MAX_TABS = 10

# Danh sách User-Agents thật
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
]

def get_random_delay():
    """Trả về delay ngẫu nhiên để giả lập hành vi người dùng thật"""
    return random.uniform(1.5, 3.5)

def setup_chrome_options():
    """Thiết lập các tùy chọn Chrome để tăng độ trust"""
    chrome_options = Options()
    
    # Cài đặt cơ bản
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    
    # Thêm User-Agent ngẫu nhiên
    chrome_options.add_argument(f'user-agent={random.choice(USER_AGENTS)}')
    
    # Tắt các tính năng automation
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Thêm các preferences để tăng độ trust
    prefs = {
        "profile.default_content_setting_values.notifications": 2,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_settings.geolocation": 2,
        "profile.managed_default_content_settings.images": 1,
        "profile.default_content_setting_values.cookies": 1
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Thêm các arguments để tăng độ trust
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-save-password-bubble")
    chrome_options.add_argument("--disable-translate")
    chrome_options.add_argument("--disable-zero-browsers-open-for-tests")
    
    return chrome_options

def simulate_human_behavior(driver):
    """Giả lập hành vi người dùng thật"""
    # Cuộn trang ngẫu nhiên
    scroll_amount = random.randint(300, 700)
    driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
    time.sleep(get_random_delay())
    
    # Cuộn lên một chút
    driver.execute_script(f"window.scrollBy(0, -{random.randint(100, 200)});")
    time.sleep(get_random_delay())
    
    # Di chuyển chuột ngẫu nhiên (giả lập)
    driver.execute_script("""
        var event = new MouseEvent('mousemove', {
            'view': window,
            'bubbles': true,
            'cancelable': true,
            'clientX': arguments[0],
            'clientY': arguments[1]
        });
        document.dispatchEvent(event);
    """, random.randint(100, 700), random.randint(100, 500))
    time.sleep(get_random_delay())

# Đọc file Excel
try:
    print(f"Đang đọc file Excel từ: {excel_path}")
    data = pd.read_excel(excel_path)
    print(f"Đã đọc file Excel thành công. Số lượng sản phẩm: {len(data)}")
except Exception as e:
    print(f"Lỗi khi đọc file Excel: {str(e)}")
    exit(1)

def fill_registration_form(driver, row, tab_index):
    try:
        # Chuyển đến tab tương ứng
        driver.switch_to.window(driver.window_handles[tab_index])
        
        # Mở URL trực tiếp từ file Excel
        url = row['product_url']
        print(f"\n{'='*50}")
        print(f"Tab {tab_index+1}: Đang xử lý sản phẩm")
        print(f"URL: {url}")
        print(f"{'='*50}")
        
        driver.get(url)
        
        # Giả lập hành vi người dùng
        simulate_human_behavior(driver)
        
        # Đợi trang load hoàn tất
        print(f"Tab {tab_index+1}: Đang đợi trang web tải hoàn tất...")
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Email']"))
        )
        
        # Điền thông tin cá nhân với delay ngẫu nhiên
        print(f"Tab {tab_index+1}: Đang điền thông tin...")
        
        # Email
        email_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Email']"))
        )
        for char in row['email']:
            email_input.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))
        time.sleep(get_random_delay())
        
        # Tên
        name_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Tên']")
        for char in row['Tên']:
            name_input.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))
        time.sleep(get_random_delay())
        
        # Họ
        lastname_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Họ']")
        for char in row['họ']:
            lastname_input.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))
        time.sleep(get_random_delay())
        
        # Số điện thoại
        phone_xpath = "//input[@placeholder='Số điện thoại' or contains(@placeholder, 'phone')]"
        phone_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, phone_xpath))
        )
        for char in str(row['số di động']):
            phone_input.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))
        time.sleep(get_random_delay())
        
        # Số serial
        serial_xpath = "//input[@placeholder='Số serial' or contains(@placeholder, 'serial')]"
        serial_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, serial_xpath))
        )
        for char in str(row['Số serial']):
            serial_input.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))
        time.sleep(get_random_delay())
        
        # Chọn ngày mua (ngày hiện tại)
        try:
            today = datetime.now()
            date_str = today.strftime("%Y-%m-%d")
            
            # Thử nhiều cách để set ngày
            try:
                # Cách 1: Tìm input type date
                date_input = driver.find_element(By.XPATH, "//input[@type='date']")
                driver.execute_script(f"arguments[0].value = '{date_str}';", date_input)
                driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", date_input)
            except:
                try:
                    # Cách 2: Tìm theo placeholder
                    date_input = driver.find_element(By.XPATH, "//input[contains(@placeholder, 'Ngày') or contains(@placeholder, 'date')]")
                    date_input.send_keys(date_str)
                except:
                    try:
                        # Cách 3: Tìm theo class có chứa date/calendar
                        date_input = driver.find_element(By.CSS_SELECTOR, "[class*='date'], [class*='calendar']")
                        date_input.send_keys(date_str)
                    except Exception as e:
                        print(f"Tab {tab_index+1}: Không thể điền ngày: {str(e)}")
            
            time.sleep(0.5)
        except Exception as e:
            print(f"Tab {tab_index+1}: Lỗi khi chọn ngày: {str(e)}")
        
        # Tích 2 checkbox bắt buộc
        try:
            # Tìm tất cả các checkbox
            checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
            
            # Tìm các label/div chứa text
            privacy_elements = driver.find_elements(By.XPATH, 
                "//*[contains(text(), 'Chính Sách Bảo Mật') or contains(text(), 'chính sách bảo mật')]")
            warranty_elements = driver.find_elements(By.XPATH, 
                "//*[contains(text(), 'Điều khoản và điều kiện bảo hành') or contains(text(), 'điều khoản bảo hành')]")
            
            # Tìm checkbox gần nhất với mỗi text
            for element in privacy_elements + warranty_elements:
                try:
                    # Tìm checkbox gần nhất
                    checkbox = element.find_element(By.XPATH, "./ancestor::div[contains(@class, 'checkbox') or contains(@class, 'form-group')]//input[@type='checkbox']")
                    if not checkbox.is_selected():
                        driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(0.3)
                except:
                    continue
            
            # Nếu không tìm thấy bằng cách trên, thử tích 2 checkbox cuối
            if len(checkboxes) >= 2:
                for checkbox in checkboxes[-2:]:
                    if not checkbox.is_selected():
                        driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(0.3)
            
        except Exception as e:
            print(f"Tab {tab_index+1}: Lỗi khi tích checkbox: {str(e)}")
        
        # Đợi người dùng xử lý captcha
        print(f"Tab {tab_index+1}: Vui lòng xử lý captcha nếu có và nhấn Enter để tiếp tục...")
        
    except Exception as e:
        print(f"Tab {tab_index+1}: Lỗi khi xử lý form: {str(e)}")

def process_batch(data_batch):
    # Thiết lập Chrome driver với các tùy chọn mới
    chrome_options = setup_chrome_options()
    
    try:
        # Khởi tạo trình duyệt
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
        # Thêm script để che giấu selenium và thêm các thuộc tính của trình duyệt thật
        driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': '''
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                });
                
                // Thêm các thuộc tính của trình duyệt thật
                Object.defineProperty(navigator, 'languages', {
                    get: () => ['vi-VN', 'vi', 'en-US', 'en']
                });
                
                Object.defineProperty(navigator, 'plugins', {
                    get: () => [1, 2, 3, 4, 5]
                });
                
                Object.defineProperty(navigator, 'platform', {
                    get: () => 'Win32'
                });
                
                // Thêm window.chrome
                window.chrome = {
                    runtime: {},
                    loadTimes: function() {},
                    csi: function() {},
                    app: {}
                };
            '''
        })
        
        # Mở các tab mới
        for i in range(len(data_batch) - 1):
            driver.execute_script("window.open('about:blank', '_blank');")
            time.sleep(get_random_delay())
        
        # Điền form cho từng tab
        for idx, row in data_batch.iterrows():
            fill_registration_form(driver, row, idx % len(data_batch))
        
        # Đợi người dùng hoàn thành captcha trên tất cả các tab
        input("\nSau khi hoàn thành tất cả captcha, nhấn Enter để tiếp tục...")
        
    except Exception as e:
        print(f"Lỗi khi xử lý batch: {str(e)}")
    finally:
        driver.quit()

def main():
    # Chia dữ liệu thành các batch
    total_records = len(data)
    batch_size = min(MAX_TABS, total_records)
    
    for i in range(0, total_records, batch_size):
        batch = data.iloc[i:i+batch_size]
        print(f"\nĐang xử lý batch {i//batch_size + 1}/{(total_records + batch_size - 1)//batch_size}")
        
        # Xử lý batch
        process_batch(batch)
        
        print(f"\nĐã hoàn thành batch {i//batch_size + 1}")
        
        # Hỏi người dùng có muốn tiếp tục batch tiếp theo không
        if i + batch_size < total_records:
            response = input("\nBạn có muốn tiếp tục batch tiếp theo không? (Y/N): ")
            if response.lower() != 'y':
                print("Đã dừng chương trình theo yêu cầu.")
                break

if __name__ == "__main__":
    main()
