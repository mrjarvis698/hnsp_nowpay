from selenium import webdriver
driver=webdriver.Chrome(executable_path="chromedriver.exe")
frames = driver.find_element_by_tag_name('iFrame')
print(len(frames))
for f in frames:
    print("Frame ID:", f.get_attribute('id'), f.get_attribute('name'))