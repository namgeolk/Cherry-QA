import pyautogui
import pydirectinput
import time

# 이미지
image = pyautogui.locateOnScreen('D.png', confidence=0.5)

# 이미지가 발견되었는지 확인
if image is not None:
    # 이미지의 중심 좌표 계산
    center_x, center_y = pyautogui.center(image)
    
    # 마우스 이동 및 클릭
    pydirectinput.moveTo(center_x, center_y,)
    pydirectinput.click()
    print(f"Clicked on the image at ({center_x}, {center_y})")
else:
    print("Image not found on the screen")

# 딜레이를 두고 반복해서 이미지 찾기
time.sleep(1)